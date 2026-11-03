/*!
 * 圣骑士环破伤害计算器 · 网页版
 * 逐条复刻 Excel「圣骑士环破伤害计算器 v1.5」（Wings 圣骑士四转环破木桩伤害）
 * 口径说明：Excel INT() 对正数等价于 Math.trunc()；攻速等级 → 延迟(ms) → 每分钟攻击次数 = 60000 / 延迟
 */
(function () {
  'use strict';

  /* ============================ 模型 ============================ */

  // 原表 G3:K10 攻速表。等级 2 最快、7 最慢（9 行延迟为空，原表 J 列为 #DIV/0!，故不提供）
  // J 列公式 = 60 / (K*100/1000) * 100 = 60000 / 延迟ms
  var SPEED_TABLE = [
    { level: 2, text: '2 · 最快', delay: 630 },
    { level: 3, text: '3 · 急速', delay: 690 },
    { level: 4, text: '4 · 更快', delay: 750 },
    { level: 5, text: '5 · 快', delay: 810 },
    { level: 6, text: '6 · 普通', delay: 840 },
    { level: 7, text: '7 · 慢', delay: 900 }
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

  var SLOT_LABELS = ['方案 1', '方案 2', '方案 3', '方案 4'];

  var WEAPON_FIELDS = [
    { key: 'str',      label: '力量',           kind: 'num', step: 1,    hint: '角色力量总和（含技能、装备、家族）' },
    { key: 'dex',      label: '敏捷',           kind: 'num', step: 1,    hint: '角色敏捷总和（含技能、装备、家族）' },
    { key: 'atk',      label: '攻击力总和',     kind: 'num', step: 1,    hint: '武器 / 手套 / 盾牌 / 护肩等合计' },
    { key: 'spd',      label: '攻速等级',       kind: 'speed',           hint: '填快速 / 急速 / 粽子 buff 后的攻速（2 最快）' },
    { key: 'mastery',  label: '熟练度',         kind: 'num', step: 0.01, min: 0, max: 1, hint: '精通点满：剑 0.6，钝器 0.75' },
    { key: 'advCoef',  label: '优势主属性系数', kind: 'num', step: 0.01, hint: '武器系数，无需改动' },
    { key: 'disCoef',  label: '劣势主属性系数', kind: 'num', step: 0.01, hint: '武器系数，无需改动' },
    { key: 'subCoef',  label: '副属性系数',     kind: 'num', step: 0.01, hint: '武器系数，无需改动' },
    { key: 'advRatio', label: '优势比例 %',     kind: 'num', step: 1,    hint: '由劈刺比决定：环破 3:2，剑填 100%，钝器填 60%' },
    { key: 'disRatio', label: '劣势比例 %',     kind: 'num', step: 1,    hint: '由劈刺比决定：环破 3:2，剑填 0%，钝器填 40%' },
    { key: 'def',      label: '怪物物理防御',   kind: 'num', step: 1,    hint: '木桩默认 0，可填怪物物理防御力' },
    { key: 'skill',    label: '技能百分比 %',   kind: 'num', step: 10,   hint: '环破点满技改 580' },
    { key: 'buff',     label: '增益百分比 %',   kind: 'num', step: 10,   hint: '技改 130 / 圣 buff 140 / 不开 buff 100' },
    { key: 'relation', label: '是否克制',       kind: 'relation',        hint: '克制 150% / 非克制 100% / 抗性 50%' }
  ];

  var WEAPON_RESULTS = [
    { key: 'advHi', label: '优势面板上限',            dec: 2 },
    { key: 'advLo', label: '优势面板下限',            dec: 2 },
    { key: 'disHi', label: '劣势面板上限',            dec: 2 },
    { key: 'disLo', label: '劣势面板下限',            dec: 2 },
    { key: 'panel', label: '面板期望值',              dec: 2, strong: true },
    { key: 'hits',  label: '每分钟攻击次数',          dec: 2 },
    { key: 'hit',   label: '单次伤害期望',            dec: 2 },
    { key: 'dpm',   label: '每分钟造成伤害（万）',    dec: 2, strong: true }
  ];

  var CONVERT_FIELDS = [
    { key: 'advCoef', label: '优势武器系数', kind: 'num', step: 0.01, hint: '武器系数，无需改动' },
    { key: 'disCoef', label: '劣势武器系数', kind: 'num', step: 0.01, hint: '武器系数，无需改动' },
    { key: 'str',     label: '力量',         kind: 'num', step: 1,    hint: '角色力量总和' },
    { key: 'dex',     label: '敏捷',         kind: 'num', step: 1,    hint: '角色敏捷总和' },
    { key: 'atk',     label: '总攻击力',     kind: 'num', step: 1,    hint: '角色攻击力总和' },
    { key: 'mastery', label: '熟练度',       kind: 'num', step: 0.01, hint: '剑 0.6，钝器 0.75' }
  ];

  var CONVERT_RESULTS = [
    { key: 'atkPerStr', label: '攻击 / 力量收益', dec: 4, strong: true },
    { key: 'strPerDex', label: '力量 / 敏捷收益', dec: 4 },
    { key: 'atkPerDex', label: '攻击 / 敏捷收益', dec: 4, strong: true }
  ];

  var DEFAULTS = {
    crit: '不计算',
    echo: '不计算',
    sword: [
      { name: '单手剑 A', str: 1200, dex: 120, atk: 200, spd: 2, mastery: 0.6, advCoef: 1,    disCoef: 1,    subCoef: 1, advRatio: 100, disRatio: 0,  def: 0, skill: 580, buff: 130, relation: '克制' },
      { name: '单手剑 B', str: 1200, dex: 120, atk: 200, spd: 2, mastery: 0.6, advCoef: 1,    disCoef: 1,    subCoef: 1, advRatio: 100, disRatio: 0,  def: 0, skill: 580, buff: 130, relation: '克制' },
      { name: '双手剑 A', str: 1400, dex: 144, atk: 250, spd: 2, mastery: 0.6, advCoef: 1.15, disCoef: 1.15, subCoef: 1, advRatio: 100, disRatio: 0,  def: 0, skill: 580, buff: 130, relation: '克制' },
      { name: '双手剑 B', str: 1400, dex: 144, atk: 200, spd: 2, mastery: 0.6, advCoef: 1.15, disCoef: 1.15, subCoef: 1, advRatio: 100, disRatio: 0,  def: 0, skill: 580, buff: 130, relation: '克制' }
    ],
    blunt: [
      { name: '单手钝器 A', str: 1300, dex: 120, atk: 200, spd: 2, mastery: 0.75, advCoef: 1.1, disCoef: 0.8,  subCoef: 1, advRatio: 60, disRatio: 40, def: 0, skill: 580, buff: 130, relation: '克制' },
      { name: '单手钝器 B', str: 1400, dex: 120, atk: 200, spd: 2, mastery: 0.75, advCoef: 1.1, disCoef: 0.8,  subCoef: 1, advRatio: 60, disRatio: 40, def: 0, skill: 580, buff: 130, relation: '克制' },
      { name: '双手钝器 A', str: 1300, dex: 120, atk: 250, spd: 2, mastery: 0.75, advCoef: 1.2, disCoef: 0.85, subCoef: 1, advRatio: 60, disRatio: 40, def: 0, skill: 580, buff: 130, relation: '克制' },
      { name: '双手钝器 B', str: 1300, dex: 120, atk: 250, spd: 2, mastery: 0.75, advCoef: 1.2, disCoef: 0.85, subCoef: 1, advRatio: 60, disRatio: 40, def: 0, skill: 580, buff: 130, relation: '克制' }
    ],
    convert: [
      { name: '双手剑',   advCoef: 1.15, disCoef: 1.15, str: 1400, dex: 144, atk: 200, mastery: 0.6  },
      { name: '单手剑',   advCoef: 1,    disCoef: 1,    str: 1100, dex: 120, atk: 200, mastery: 0.6  },
      { name: '双手钝器', advCoef: 1.2,  disCoef: 0.85, str: 1400, dex: 144, atk: 200, mastery: 0.75 },
      { name: '单手钝器', advCoef: 1.1,  disCoef: 0.8,  str: 1000, dex: 100, atk: 200, mastery: 0.75 }
    ]
  };

  /* ------------------------- 计算 ------------------------- */

  function optBy(list, text, fallback) {
    for (var i = 0; i < list.length; i++) if (list[i].text === text) return list[i];
    return list[fallback || 0];
  }

  function delayOf(level) {
    for (var i = 0; i < SPEED_TABLE.length; i++) if (SPEED_TABLE[i].level === level) return SPEED_TABLE[i].delay;
    return 0;
  }

  // 单方案计算：X14..X25
  function computeColumn(c, critLabel, echoLabel) {
    var crit = optBy(CRIT_OPTIONS, critLabel);
    var echo = optBy(ECHO_OPTIONS, echoLabel);
    var rel = optBy(RELATION_OPTIONS, c.relation);

    // X4 经回声加成后取整，再除以 100（原表 INT(X4*(1+I16))/100）
    var atkEff = Math.trunc(c.atk * (1 + echo.bonus)) / 100;

    var advHi = (c.str * c.advCoef * 4 + c.dex * c.subCoef) * atkEff;
    var advLo = (c.str * c.advCoef * 4 * 0.9 * c.mastery + c.dex * c.subCoef) * atkEff;
    var disHi = (c.str * c.disCoef * 4 + c.dex * c.subCoef) * atkEff;
    var disLo = (c.str * c.disCoef * 4 * 0.9 * c.mastery + c.dex * c.subCoef) * atkEff;

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
    var k = 0.9 * c.mastery + 1;
    var den = 12 * c.atk * c.advCoef * k + 8 * c.atk * c.disCoef * k;
    var num = (c.atk + 1) * (12 * c.str * c.advCoef * k + 8 * c.str * c.disCoef * k + 10 * c.dex) - 10 * c.atk * c.dex;
    var atkPerStr = (den === 0 ? NaN : num / den) - c.str;
    var strPerDex = (1 + 0.9 * c.mastery) * (1.2 * c.advCoef + 0.8 * c.disCoef);
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

  var STORAGE_KEY = 'pldc-state-v1';

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

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

  function pickText(v, list, d) {
    for (var i = 0; i < list.length; i++) if (list[i].text === v) return v;
    return d;
  }

  function normalizeSlot(src, base, isConvert) {
    var o = clone(base);
    if (!src || typeof src !== 'object') return o;
    for (var k in base) {
      if (!base.hasOwnProperty(k)) continue;
      if (k === 'name') { o.name = typeof src.name === 'string' && src.name.trim() ? src.name.slice(0, 12) : base.name; continue; }
      if (k === 'relation') { o.relation = pickRelation(src.relation, base.relation); continue; }
      if (k === 'spd' && !isConvert) { o.spd = pickSpd(src.spd, base.spd); continue; }
      o[k] = num(src[k], base[k]);
    }
    return o;
  }

  function normalize(raw) {
    var d = clone(DEFAULTS);
    if (!raw || typeof raw !== 'object') return d;
    var out = { crit: pickText(raw.crit, CRIT_OPTIONS, DEFAULTS.crit), echo: pickText(raw.echo, ECHO_OPTIONS, DEFAULTS.echo) };
    ['sword', 'blunt', 'convert'].forEach(function (mode) {
      out[mode] = [];
      for (var i = 0; i < 4; i++) {
        out[mode].push(normalizeSlot(raw[mode] && raw[mode][i], DEFAULTS[mode][i], mode === 'convert'));
      }
    });
    return out;
  }

  function encodeState(s) {
    var json = JSON.stringify(s);
    return btoa(unescape(encodeURIComponent(json))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  function decodeState(str) {
    var b = str.replace(/-/g, '+').replace(/_/g, '/');
    while (b.length % 4) b += '=';
    return JSON.parse(decodeURIComponent(escape(atob(b))));
  }

  function readHashState() {
    var h = location.hash || '';
    var i = h.indexOf('d=');
    if (i < 0) return null;
    try { return decodeState(h.slice(i + 2)); } catch (e) { return null; }
  }

  function loadState() {
    var fromHash = readHashState();
    if (fromHash) return normalize(fromHash);
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return normalize(JSON.parse(raw));
    } catch (e) { /* localStorage 不可用则忽略 */ }
    return normalize(null);
  }

  var saveTimer = null;
  function saveState() {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(S)); } catch (e) { /* 忽略配额错误 */ }
    }, 300);
  }

  /* ------------------------- 视图 ------------------------- */

  var S = null;
  var root = null;

  function speedOptions(cur) {
    var out = '';
    for (var i = 0; i < SPEED_TABLE.length; i++) {
      var s = SPEED_TABLE[i];
      out += '<option value="' + s.level + '"' + (s.level === cur ? ' selected' : '') + '>' +
        s.text + ' · ' + s.delay + 'ms</option>';
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

  function cellControl(mode, col, f, slot) {
    var base = ' data-mode="' + mode + '" data-col="' + col + '" data-key="' + f.key + '"';
    if (f.kind === 'speed') {
      return '<select class="pldc-input pldc-select"' + base + '>' + speedOptions(slot.spd) + '</select>';
    }
    if (f.kind === 'relation') {
      return '<select class="pldc-input pldc-select"' + base + '>' + relationOptions(slot.relation) + '</select>';
    }
    var attrs = ' type="number" inputmode="decimal" step="' + (f.step || 1) + '"' +
      (typeof f.min === 'number' ? ' min="' + f.min + '"' : '') +
      (typeof f.max === 'number' ? ' max="' + f.max + '"' : '');
    return '<input class="pldc-input"' + base + attrs + ' value="' + slot[f.key] + '">';
  }

  function weaponPanel(mode) {
    var slots = S[mode];
    var html = '<section class="pldc-panel" data-panel="' + mode + '" hidden>' +
      '<div class="pldc-cards" data-cards="' + mode + '"></div>' +
      '<p class="pldc-verdict" data-verdict="' + mode + '"></p>' +
      '<div class="pldc-tablewrap"><table class="pldc-table">' +
      '<thead><tr><th class="pldc-rowhead" scope="col">参数</th>';

    for (var i = 0; i < 4; i++) {
      html += '<th scope="col" class="pldc-colhead">' +
        '<input class="pldc-name" data-mode="' + mode + '" data-col="' + i + '" data-key="name" ' +
        'type="text" maxlength="12" spellcheck="false" value="' + esc(slots[i].name) + '" title="点击修改方案名称">' +
        '</th>';
    }
    html += '</tr></thead><tbody>';

    for (var f = 0; f < WEAPON_FIELDS.length; f++) {
      var fld = WEAPON_FIELDS[f];
      html += '<tr><th scope="row" class="pldc-rowhead" title="' + esc(fld.hint) + '">' +
        esc(fld.label) + '<span class="pldc-tip" aria-hidden="true">?</span></th>';
      for (var c = 0; c < 4; c++) {
        html += '<td class="pldc-cell">' + cellControl(mode, c, fld, slots[c]) +
          '<button type="button" class="pldc-sync" data-act="sync" data-mode="' + mode + '" data-col="' + c +
          '" data-key="' + fld.key + '" title="把该值同步到本表另外三列" aria-label="同步">⇊</button></td>';
      }
      html += '</tr>';
    }

    html += '<tr class="pldc-divider"><th colspan="5" scope="colgroup">计算结果</th></tr>';
    for (var r = 0; r < WEAPON_RESULTS.length; r++) {
      var res = WEAPON_RESULTS[r];
      html += '<tr class="' + (res.strong ? 'pldc-row-strong' : '') + '">' +
        '<th scope="row" class="pldc-rowhead">' + esc(res.label) + '</th>';
      for (var c2 = 0; c2 < 4; c2++) {
        html += '<td class="pldc-cell pldc-out" data-out="' + mode + '" data-col="' + c2 +
          '" data-key="' + res.key + '"></td>';
      }
      html += '</tr>';
    }
    html += '</tbody></table></div>' +
      '<p class="pldc-note">灰色小字为可编辑参数；把鼠标移到参数名上可看到原表说明。「⇊」按钮把该格数值同步到本表另外三列。' +
      '比例按 <b>%</b> 填写（原表 0.6 = 本页 60%），技能 / 增益百分比同理（原表 5.8 = 本页 580）。</p>' +
      '</section>';
    return html;
  }

  function convertPanel() {
    var slots = S.convert;
    var html = '<section class="pldc-panel" data-panel="convert" hidden>' +
      '<p class="pldc-verdict">把四把武器的力量 / 敏捷 / 攻击力填进去，得到「1 攻击 ≈ X 点力量」的等效换算。' +
      '本表算整体输出提升，不是面板最高值；未包含英雄回声（回声请在环破页勾选）。</p>' +
      '<div class="pldc-tablewrap"><table class="pldc-table">' +
      '<thead><tr><th class="pldc-rowhead" scope="col">参数</th>';
    for (var i = 0; i < 4; i++) {
      html += '<th scope="col" class="pldc-colhead">' +
        '<input class="pldc-name" data-mode="convert" data-col="' + i + '" data-key="name" ' +
        'type="text" maxlength="12" spellcheck="false" value="' + esc(slots[i].name) + '" title="点击修改武器名称">' +
        '</th>';
    }
    html += '</tr></thead><tbody>';

    for (var f = 0; f < CONVERT_FIELDS.length; f++) {
      var fld = CONVERT_FIELDS[f];
      html += '<tr><th scope="row" class="pldc-rowhead" title="' + esc(fld.hint) + '">' +
        esc(fld.label) + '<span class="pldc-tip" aria-hidden="true">?</span></th>';
      for (var c = 0; c < 4; c++) {
        html += '<td class="pldc-cell">' + cellControl('convert', c, fld, slots[c]) +
          '<button type="button" class="pldc-sync" data-act="sync" data-mode="convert" data-col="' + c +
          '" data-key="' + fld.key + '" title="把该值同步到本表另外三列" aria-label="同步">⇊</button></td>';
      }
      html += '</tr>';
    }

    html += '<tr class="pldc-divider"><th colspan="5" scope="colgroup">换算结果</th></tr>';
    for (var r = 0; r < CONVERT_RESULTS.length; r++) {
      var res = CONVERT_RESULTS[r];
      html += '<tr class="' + (res.strong ? 'pldc-row-strong' : '') + '">' +
        '<th scope="row" class="pldc-rowhead">' + esc(res.label) + '</th>';
      for (var c2 = 0; c2 < 4; c2++) {
        html += '<td class="pldc-cell pldc-out" data-out="convert" data-col="' + c2 +
          '" data-key="' + res.key + '"></td>';
      }
      html += '</tr>';
    }

    html += '</tbody></table></div>' +
      '<p class="pldc-note">读法：攻击 / 力量收益 = 7.2 表示 1 点攻击力与 7.2 点力量的输出提升相同；' +
      '攻击 / 敏捷收益 = 25.5 表示 1 点攻击与 25.5 点敏捷等价。' +
      '未计算英雄回声，也未计算压制术与伤害上限。</p>' +
      '</section>';
    return html;
  }

  function toolbarHtml() {
    var crit = '', echo = '';
    for (var i = 0; i < CRIT_OPTIONS.length; i++) {
      crit += '<option value="' + CRIT_OPTIONS[i].text + '"' + (CRIT_OPTIONS[i].text === S.crit ? ' selected' : '') + '>' +
        CRIT_OPTIONS[i].text + '</option>';
    }
    for (var j = 0; j < ECHO_OPTIONS.length; j++) {
      echo += '<option value="' + ECHO_OPTIONS[j].text + '"' + (ECHO_OPTIONS[j].text === S.echo ? ' selected' : '') + '>' +
        ECHO_OPTIONS[j].text + '</option>';
    }
    return '<label class="pldc-field" title="I12：计算暴击增益的方式，暴击率与暴击伤害自动跟随">' +
      '<span>暴击增益</span><select class="pldc-input pldc-select" data-global="crit">' + crit + '</select></label>' +
      '<label class="pldc-field" title="I15：英雄回声，攻击力按 INT(攻击力*1.04) 计算">' +
      '<span>英雄回声</span><select class="pldc-input pldc-select" data-global="echo">' + echo + '</select></label>' +
      '<span class="pldc-grow"></span>' +
      '<button type="button" class="pldc-btn" data-act="share">复制分享链接</button>' +
      '<button type="button" class="pldc-btn pldc-btn-ghost" data-act="reset">恢复默认</button>';
  }

  var TABS = [
    { key: 'sword',   label: '剑 · 环破' },
    { key: 'blunt',   label: '钝器 · 环破' },
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

  function resultsOf(mode) {
    var out = [];
    for (var i = 0; i < 4; i++) out.push(computeColumn(S[mode][i], S.crit, S.echo));
    return out;
  }

  function setText(node, text) { if (node && node.textContent !== text) node.textContent = text; }

  function refreshWeapon(mode) {
    var res = resultsOf(mode);
    var best = -Infinity, worst = Infinity;
    for (var i = 0; i < 4; i++) {
      if (res[i].dpm > best) best = res[i].dpm;
      if (res[i].dpm < worst) worst = res[i].dpm;
    }

    var cells = root.querySelectorAll('[data-out="' + mode + '"]');
    for (var n = 0; n < cells.length; n++) {
      var cell = cells[n];
      var col = parseInt(cell.getAttribute('data-col'), 10);
      var key = cell.getAttribute('data-key');
      var spec = null;
      for (var q = 0; q < WEAPON_RESULTS.length; q++) if (WEAPON_RESULTS[q].key === key) spec = WEAPON_RESULTS[q];
      if (!spec) continue;
      setText(cell, fmtNum(res[col][key], spec.dec));
      cell.classList.toggle('pldc-best', key === 'dpm' && res[col].dpm === best && best > 0);
    }

    var cards = [];
    for (var c = 0; c < 4; c++) {
      var r = res[c];
      var isBest = r.dpm === best && best > 0;
      var delta = best > 0 && r.dpm > 0 && r.dpm < best ? '−' + fmtNum((1 - r.dpm / best) * 100, 1) + '%' : '';
      cards.push('<div class="pldc-card' + (isBest ? ' pldc-card-best' : '') + '">' +
        '<div class="pldc-card-top"><span class="pldc-card-name">' + esc(S[mode][c].name) + '</span>' +
        (isBest ? '<span class="pldc-pill">最高</span>' : (delta ? '<span class="pldc-pill pldc-pill-dim">' + delta + '</span>' : '')) +
        '</div>' +
        '<div class="pldc-card-main">' + fmtNum(r.dpm, 2) + '<small>万 / 分钟</small></div>' +
        '<div class="pldc-card-sub">单次 ' + fmtNum(r.hit, 2) + ' · 面板 ' + fmtNum(r.panel, 2) + '</div>' +
        '</div>');
    }
    var cardBox = root.querySelector('[data-cards="' + mode + '"]');
    if (cardBox) cardBox.innerHTML = cards.join('');

    var verdict = root.querySelector('[data-verdict="' + mode + '"]');
    if (verdict) {
      var bestIdx = res.map(function (x) { return x.dpm; }).indexOf(best);
      var gap = worst > 0 ? (1 - worst / best) * 100 : 0;
      var warn = '';
      for (var w = 0; w < 4; w++) {
        if (Math.abs(res[w].ratioSum - 100) > 0.001) {
          warn = '<span class="pldc-warn">注意：' + esc(S[mode][w].name) + ' 的优势 + 劣势比例为 ' +
            fmtNum(res[w].ratioSum, 1) + '%，不等于 100%</span>';
          break;
        }
      }
      verdict.innerHTML = '本表最高：<b>' + esc(S[mode][bestIdx].name) + '</b> · ' + fmtNum(best, 2) +
        ' 万 / 分钟' + (gap > 0.05 ? '，最低方案比它低 ' + fmtNum(gap, 1) + '%' : '') + '。' + warn;
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
    refreshWeapon('sword');
    refreshWeapon('blunt');
    refreshConvert();
  }

  /* ------------------------- 交互 ------------------------- */

  function readControlValue(node) {
    if (node.tagName === 'SELECT') return node.value;
    if (node.getAttribute('data-key') === 'name') return node.value;
    if (node.value === '') return '';
    return node.value;
  }

  function writeSlot(mode, col, key, raw) {
    var slot = S[mode][col];
    if (key === 'name') { slot.name = String(raw).slice(0, 12) || SLOT_LABELS[col]; return; }
    if (key === 'relation') { slot.relation = pickRelation(raw, slot.relation); return; }
    if (key === 'spd' && mode !== 'convert') { slot.spd = pickSpd(raw, slot.spd); return; }
    slot[key] = num(raw, 0);
  }

  function onInput(e) {
    var t = e.target;
    var g = t.getAttribute('data-global');
    if (g) {
      if (g === 'crit') S.crit = pickText(t.value, CRIT_OPTIONS, S.crit);
      if (g === 'echo') S.echo = pickText(t.value, ECHO_OPTIONS, S.echo);
      refreshAll();
      saveState();
      return;
    }
    var mode = t.getAttribute('data-mode');
    var col = t.getAttribute('data-col');
    var key = t.getAttribute('data-key');
    if (!mode || col === null || !key) return;
    writeSlot(mode, parseInt(col, 10), key, readControlValue(t));
    if (mode === 'convert') refreshConvert(); else refreshWeapon(mode);
    saveState();
  }

  function onBlur(e) {
    var t = e.target;
    if (!t.getAttribute || !t.getAttribute('data-key')) return;
    if (t.getAttribute('data-key') === 'name') return;
    if (t.tagName === 'SELECT') return;
    var mode = t.getAttribute('data-mode');
    var col = parseInt(t.getAttribute('data-col'), 10);
    var key = t.getAttribute('data-key');
    if (!mode || isNaN(col)) return;
    var val = S[mode][col][key];
    if (t.value !== String(val)) t.value = String(val);
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
      navigator.clipboard.writeText(text).then(function () { toast('链接已复制'); }, function () { toast('复制失败，请手动复制地址栏'); });
      return;
    }
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', 'readonly');
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); toast('链接已复制'); } catch (e) { toast('复制失败，请手动复制地址栏'); }
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
    if (bar) bar.hidden = key === 'convert' || key === 'guide';
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
    if (act === 'sync') {
      var mode = btn.getAttribute('data-mode');
      var col = parseInt(btn.getAttribute('data-col'), 10);
      var key = btn.getAttribute('data-key');
      var val = S[mode][col][key];
      for (var i = 0; i < 4; i++) {
        if (i === col) continue;
        writeSlot(mode, i, key, val);
        var input = root.querySelector('[data-mode="' + mode + '"][data-col="' + i + '"][data-key="' + key + '"]');
        if (input) input.value = val;
      }
      if (mode === 'convert') refreshConvert(); else refreshWeapon(mode);
      saveState();
      toast('已同步到本表四列');
      return;
    }
  }

  function build() {
    var bar = document.getElementById('pldc-toolbar');
    if (bar) bar.innerHTML = toolbarHtml();
    var tabs = document.getElementById('pldc-tabs');
    if (tabs) tabs.innerHTML = tabsHtml();
    var panels = document.getElementById('pldc-panels');
    if (panels) panels.innerHTML = weaponPanel('sword') + weaponPanel('blunt') + convertPanel();
  }

  function rebuild() {
    build();
    refreshAll();
  }

  function init() {
    root = document.getElementById('pldc');
    if (!root) return;
    S = loadState();
    build();
    root.addEventListener('input', onInput);
    root.addEventListener('change', onInput);
    root.addEventListener('blur', onBlur, true);
    root.addEventListener('click', onClick);
    refreshAll();
    setTab('sword');
    root.setAttribute('data-ready', '1');
  }

  /* ------------------------- 导出 ------------------------- */

  var API = {
    computeColumn: computeColumn,
    computeConvert: computeConvert,
    normalize: normalize,
    DEFAULTS: DEFAULTS,
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
