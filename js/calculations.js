/**
 * calculations.js — محرك الحساب.
 * يطبق نفس القواعد التي تُكتب كمعادلات في Excel، ويُستخدم للعرض والمعاينة
 * ولتعبئة "النتيجة المخزنة" بجوار كل معادلة. لا يُصدَّر أي ناتج منه كقيمة ثابتة.
 *
 *   ساعات المهمة     = المرات السنوية × التكرار × المدة ÷ 60
 *   إجمالي الساعات   = Σ ساعات المهام
 *   الاحتياج الدقيق   = إجمالي الساعات ÷ ساعات العمل السنوية
 *   الاحتياج المحسوب = تقريب(الاحتياج الدقيق)  ← 0.5 فأعلى للأعلى
 *   الفجوة           = الاحتياج المحسوب − العدد الفعلي
 *   الحالة           = عجز (>0) | فائض (<0) | متوازن (=0)
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});
  var U = WL.utils;

  function findFrequency(cfg, key) {
    for (var i = 0; i < cfg.frequencies.length; i++) {
      if (cfg.frequencies[i].key === key) return cfg.frequencies[i];
    }
    return null;
  }

  /** @returns {number|null} null إذا كانت بيانات المهمة غير مكتملة */
  function taskHours(task, cfg) {
    var freq = findFrequency(cfg, task.frequencyKey);
    var rep = U.parseNumber(task.repetitions);
    var dur = U.parseNumber(task.durationMinutes);
    if (!freq || !U.isValidNumber(rep) || !U.isValidNumber(dur)) return null;
    return (freq.perYear * rep * dur) / cfg.minutesPerHour;
  }

  function statusOf(gap, cfg) {
    if (!U.isValidNumber(gap)) return { key: null, label: '—' };
    if (gap > 0) return { key: 'deficit', label: cfg.status.deficit };
    if (gap < 0) return { key: 'surplus', label: cfg.status.surplus };
    return { key: 'balanced', label: cfg.status.balanced };
  }

  function analyze(project, cfg) {
    var rows = project.tasks.map(function (t) {
      var freq = findFrequency(cfg, t.frequencyKey);
      return {
        id: t.id,
        title: t.title,
        frequencyKey: t.frequencyKey,
        frequencyLabel: freq ? freq.label : '',
        perYear: freq ? freq.perYear : null,
        repetitions: U.parseNumber(t.repetitions),
        durationMinutes: U.parseNumber(t.durationMinutes),
        hours: taskHours(t, cfg),
        share: null
      };
    });

    var totalHours = rows.reduce(function (sum, r) {
      return sum + (r.hours === null ? 0 : r.hours);
    }, 0);

    rows.forEach(function (r) {
      r.share = r.hours === null ? null : (totalHours === 0 ? 0 : r.hours / totalHours);
    });

    var byFrequency = cfg.frequencies.map(function (f) {
      var matched = rows.filter(function (r) { return r.frequencyKey === f.key; });
      var hours = matched.reduce(function (s, r) { return s + (r.hours || 0); }, 0);
      return {
        key: f.key,
        label: f.label,
        perYear: f.perYear,
        count: matched.length,
        hours: hours,
        share: totalHours === 0 ? 0 : hours / totalHours
      };
    });

    var annual = cfg.annualWorkHours;
    var exactNeed = annual > 0 ? totalHours / annual : 0;
    var calcNeed = U.roundHalfUp(exactNeed, cfg.needRounding.digits);

    var actual = U.parseNumber(project.org.actualCount);
    var gap = U.isValidNumber(actual) ? calcNeed - actual : null;
    var status = statusOf(gap, cfg);

    return {
      tasks: rows,
      taskCount: rows.length,
      totalHours: totalHours,
      annualWorkHours: annual,
      exactNeed: exactNeed,
      calcNeed: calcNeed,
      actualCount: U.isValidNumber(actual) ? actual : null,
      gap: gap,
      statusKey: status.key,
      status: status.label,
      byFrequency: byFrequency
    };
  }

  WL.calc = {
    findFrequency: findFrequency,
    taskHours: taskHours,
    statusOf: statusOf,
    analyze: analyze
  };
})(typeof window !== 'undefined' ? window : globalThis);
