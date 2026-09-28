/**
 * excel-export.js — محرك Excel.
 * يبني ملف XLSX حقيقيًا: البيانات المدخلة كقيم، وكل النتائج كمعادلات Excel فعلية
 * تعتمد على أسماء معرّفة (Named Ranges) وعلى ورقة الإعدادات.
 * القيمة المحسوبة في JavaScript تُخزن فقط كـ "نتيجة مخزنة" للمعادلة، ويعيد Excel الحساب عند الفتح.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});
  var U = WL.utils;

  // ───────────── أدوات مساعدة ─────────────

  function f(formula, result) {
    var v = { formula: formula };
    if (result !== undefined && result !== null && !(typeof result === 'number' && !isFinite(result))) {
      v.result = result;
    }
    return v;
  }

  function sheetRef(ctx, key, addr) {
    return "'" + ctx.cfg.excel.sheetNames[key] + "'!" + addr;
  }

  function nameRange(ws, col, fromRow, toRow, name) {
    for (var r = fromRow; r <= toRow; r++) ws.getCell(col + r).addName(name);
  }

  function mergeStyled(ws, range, value, styler) {
    ws.mergeCells(range);
    var cell = ws.getCell(range.split(':')[0]);
    if (value !== undefined) cell.value = value;
    if (styler) styler(cell);
    return cell;
  }

  function needFormat(digits) {
    return digits > 0 ? '#,##0.' + new Array(digits + 1).join('0') : '#,##0';
  }

  function gapFormat(digits) {
    var n = needFormat(digits);
    return '+' + n + ';-' + n + ';0';
  }

  function blankToNull(v) {
    return v === undefined || v === null || String(v).trim() === '' ? null : v;
  }

  function estimateRowHeight(text, colWidth) {
    var len = String(text || '').length;
    var perLine = Math.max(10, Math.floor(colWidth * 1.15));
    var lines = Math.max(1, Math.ceil(len / perLine));
    return Math.max(22, lines * 16 + 6);
  }

  function titleBlock(ctx, ws, lastCol, title, withEntity) {
    var T = ctx.T;
    mergeStyled(ws, 'A1:' + lastCol + '1', title, T.style.title);
    ws.getRow(1).height = 34;
    if (withEntity) {
      var N = T.NAMES;
      mergeStyled(ws, 'A2:' + lastCol + '2',
        f('"الجهة: "&' + N.entity, 'الجهة: ' + ctx.project.org.entity),
        T.style.subtitle);
      ws.getRow(2).height = 22;
    }
  }

  function numberValidation(kind, operator, value, message) {
    return {
      type: kind, operator: operator, allowBlank: true, formulae: [value],
      showErrorMessage: true, errorStyle: 'stop',
      errorTitle: 'قيمة غير صحيحة', error: message
    };
  }

  // ───────────── ورقة الإعدادات ─────────────

  function buildSettings(ctx) {
    var ws = ctx.sheets.settings, T = ctx.T, N = T.NAMES, L = T.LAYOUT.settings, cfg = ctx.cfg;
    T.setColumnWidths(ws, L.columns);
    titleBlock(ctx, ws, 'C', 'الإعدادات والقيم المرجعية', false);
    mergeStyled(ws, 'A2:C2', 'جميع معادلات الملف تعتمد على القيم في هذه الورقة. الخلية الملونة بالأصفر الفاتح قابلة للتعديل.', T.style.note);

    ['الإعداد', 'القيمة', 'ملاحظة'].forEach(function (h, i) {
      var c = ws.getRow(3).getCell(i + 1); c.value = h; T.style.header(c);
    });

    var r = L.annualHoursRow;
    ws.getCell('A' + r).value = 'ساعات العمل الفعلية السنوية للموظف'; T.style.label(ws.getCell('A' + r));
    var ah = ws.getCell('B' + r);
    ah.value = cfg.annualWorkHours;
    T.style.input(ah, T.NUM.integer, 'center');
    ah.addName(N.annualHours);
    ah.dataValidation = numberValidation('decimal', 'greaterThan', 0, 'ساعات العمل السنوية يجب أن تكون رقمًا أكبر من صفر.');
    ws.getCell('C' + r).value = 'يستخدم لحساب الاحتياج: إجمالي الساعات ÷ هذه القيمة';
    T.style.note(ws.getCell('C' + r));

    r = L.needDecimalsRow;
    ws.getCell('A' + r).value = 'منازل تقريب الاحتياج المحسوب'; T.style.label(ws.getCell('A' + r));
    var nd = ws.getCell('B' + r);
    nd.value = cfg.needRounding.digits;
    T.style.formula(nd, T.NUM.integer);
    nd.addName(N.needDecimals);
    ws.getCell('C' + r).value = 'تقريب لأقرب عدد صحيح: 0.5 فأعلى يُجبر للأعلى (1.5 ← 2 ، 1.3 ← 1)';
    T.style.note(ws.getCell('C' + r));

    var hr = L.freqHeaderRow;
    mergeStyled(ws, 'A' + (hr - 1) + ':C' + (hr - 1), 'جدول الترددات السنوية (قيم معتمدة)', T.style.section);

    ws.addTable({
      name: T.LAYOUT.settingsFreqTableName,
      ref: 'A' + hr,
      headerRow: true,
      totalsRow: false,
      style: { theme: 'TableStyleLight1', showRowStripes: false },
      columns: [{ name: 'التردد', filterButton: false }, { name: 'المرات السنوية', filterButton: false }],
      rows: cfg.frequencies.map(function (fr) { return [fr.label, fr.perYear]; })
    });

    var first = hr + 1, last = hr + cfg.frequencies.length;
    [ws.getCell('A' + hr), ws.getCell('B' + hr)].forEach(T.style.header);
    for (var i = first; i <= last; i++) {
      T.style.formula(ws.getCell('A' + i), null, { align: 'right' });
      T.style.formula(ws.getCell('B' + i), T.NUM.integer);
      ws.getRow(i).height = 22;
    }
    [3, L.annualHoursRow, L.needDecimalsRow, hr - 1, hr].forEach(function (n) { ws.getRow(n).height = 24; });
    nameRange(ws, 'A', first, last, N.freqNames);
    nameRange(ws, 'B', first, last, N.freqValues);

    ctx.freqRows = { first: first, last: last };
    T.applyPrint(ws, { paper: 'A4', landscape: false, printArea: 'A1:C' + last, reportTitle: cfg.app.reportTitle, dateText: ctx.dateText });
  }

  // ───────────── ورقة تحليل المهام ─────────────

  function buildTasks(ctx) {
    var ws = ctx.sheets.tasks, T = ctx.T, N = T.NAMES, L = T.LAYOUT.tasks, cfg = ctx.cfg, a = ctx.analysis;
    var cols = L.columns;
    T.setColumnWidths(ws, cols.map(function (c) { return c.width; }));
    titleBlock(ctx, ws, 'H', 'تحليل المهام', true);

    var hr = L.headerRow;
    var first = hr + 1;
    var count = a.tasks.length + cfg.excel.spareTaskRows;
    var last = hr + count;

    var rows = [];
    for (var i = 0; i < count; i++) {
      var r = first + i;
      var t = a.tasks[i];
      var noF = 'IF(B' + r + '="","",ROW()-' + hr + ')';
      var perYearF = 'IF(C' + r + '="","",IF(COUNTIF(' + N.freqNames + ',C' + r + ')=0,"",INDEX(' + N.freqValues + ',MATCH(C' + r + ',' + N.freqNames + ',0))))';
      var hoursF = 'IF(OR(D' + r + '="",E' + r + '="",F' + r + '=""),"",D' + r + '*E' + r + '*F' + r + '/' + cfg.minutesPerHour + ')';
      var shareF = 'IF(OR(G' + r + '="",' + N.totalHours + '=0),"",G' + r + '/' + N.totalHours + ')';
      if (t) {
        rows.push([
          f(noF, i + 1),
          t.title,
          t.frequencyLabel,
          f(perYearF, t.perYear),
          t.repetitions,
          t.durationMinutes,
          f(hoursF, t.hours),
          f(shareF, t.share)
        ]);
      } else {
        rows.push([f(noF, ''), null, null, f(perYearF, ''), null, null, f(hoursF, ''), f(shareF, '')]);
      }
    }

    ws.addTable({
      name: L.tableName,
      ref: 'A' + hr,
      headerRow: true,
      totalsRow: false,
      style: { theme: 'TableStyleLight1', showRowStripes: false },
      columns: cols.map(function (c) { return { name: c.header, filterButton: true }; }),
      rows: rows
    });

    cols.forEach(function (c, i) { T.style.header(ws.getRow(hr).getCell(i + 1)); });
    ws.getRow(hr).height = 30;

    var freqError = 'يرجى اختيار التردد من القائمة المنسدلة.';
    for (var k = 0; k < count; k++) {
      var rowNum = first + k;
      var row = ws.getRow(rowNum);
      var stripe = k % 2 === 1 ? T.COLORS.stripe : T.COLORS.white;
      T.style.formula(row.getCell(1), null, { fill: stripe });
      T.style.input(row.getCell(2));
      T.style.input(row.getCell(3), null, 'center');
      T.style.formula(row.getCell(4), T.NUM.integer, { fill: stripe });
      T.style.input(row.getCell(5), T.NUM.general, 'center');
      T.style.input(row.getCell(6), T.NUM.general, 'center');
      T.style.formula(row.getCell(7), T.NUM.hours, { fill: stripe, bold: true });
      T.style.formula(row.getCell(8), T.NUM.percent, { fill: stripe });

      row.getCell(3).dataValidation = {
        type: 'list', allowBlank: true, formulae: [N.freqNames],
        showErrorMessage: true, errorStyle: 'stop', errorTitle: 'تردد غير صحيح', error: freqError,
        showInputMessage: true, promptTitle: 'التردد', prompt: 'اختر التردد من القائمة'
      };
      row.getCell(5).dataValidation = numberValidation('decimal', 'greaterThan', 0, 'يجب أن يكون عدد التكرارات رقمًا أكبر من صفر.');
      row.getCell(6).dataValidation = numberValidation('decimal', 'greaterThan', 0, 'يجب أن تكون المدة بالدقائق رقمًا أكبر من صفر.');

      var task = a.tasks[k];
      row.height = task ? estimateRowHeight(task.title, cols[1].width) : 20;
    }

    nameRange(ws, 'B', first, last, N.taskTitles);
    nameRange(ws, 'C', first, last, N.taskFreq);
    nameRange(ws, 'G', first, last, N.taskHours);

    // تنبيهات داخل Excel: تردد غير معروف، أو مهمة بلا تكرار/مدة
    T.addWarningRule(ws, 'C' + first + ':C' + last, 'AND($C' + first + '<>"",$D' + first + '="")');
    T.addWarningRule(ws, 'C' + first + ':C' + last, 'AND($B' + first + '<>"",$C' + first + '="")');
    T.addWarningRule(ws, 'E' + first + ':F' + last, 'AND($B' + first + '<>"",E' + first + '="")');

    // الإجماليات أعلى الجدول (تبقى ظاهرة مع تجميد الأجزاء)
    var tr = L.totalRow, cr = L.countRow;
    mergeStyled(ws, 'A' + tr + ':F' + tr, 'إجمالي ساعات العمل السنوية', T.style.label);
    var total = ws.getCell('G' + tr);
    total.value = f('SUM(' + N.taskHours + ')', a.totalHours);
    T.style.total(total, T.NUM.hours);
    total.addName(N.totalHours);
    var shareTotal = ws.getCell('H' + tr);
    shareTotal.value = f('SUM(H' + first + ':H' + last + ')', a.totalHours > 0 ? 1 : 0);
    T.style.total(shareTotal, T.NUM.percent);

    mergeStyled(ws, 'A' + cr + ':F' + cr, 'عدد المهام', T.style.label);
    var cnt = ws.getCell('G' + cr);
    cnt.value = f('COUNTA(' + N.taskTitles + ')', a.taskCount);
    T.style.total(cnt, T.NUM.integer);
    cnt.addName(N.taskCount);
    T.style.total(ws.getCell('H' + cr));
    ws.getRow(tr).height = 24;
    ws.getRow(cr).height = 24;

    mergeStyled(ws, 'A' + L.noteRow + ':H' + L.noteRow,
      'الخلايا الملونة بالأصفر الفاتح للإدخال. اختر التردد من القائمة وتُحسب بقية الأعمدة تلقائيًا. يوجد ' +
      cfg.excel.spareTaskRows + ' صفًا فارغًا جاهزًا أسفل المهام لإضافة مهام جديدة، ويتسع نطاق الطباعة تلقائيًا ليشملها.', T.style.note);
    ws.getRow(L.noteRow).height = 20;

    var paper = a.tasks.length + hr > cfg.excel.a3RowThreshold ? 'A3' : 'A4';
    T.applyPrint(ws, {
      paper: paper, landscape: true, printArea: 'A1:H' + (hr + a.tasks.length),
      titleRows: hr + ':' + hr, reportTitle: cfg.app.reportTitle, dateText: ctx.dateText
    });

    // نطاق طباعة ديناميكي: حتى آخر صف يحتوي مهمة (بدل طباعة الصفوف الاحتياطية الفارغة)
    var sn = "'" + cfg.excel.sheetNames.tasks + "'!";
    var titles = sn + '$B$' + first + ':$B$' + last;
    ctx.dynamicPrintAreas.push({
      sheetKey: 'tasks',
      formula: 'OFFSET(' + sn + '$A$1,0,0,IF(COUNTA(' + titles + ')=0,' + hr + ',LOOKUP(2,1/(' + titles + '<>""),ROW(' + titles + '))),' + cols.length + ')'
    });
    ctx.taskRows = { first: first, last: last };
  }

  // ───────────── ورقة البيانات الأساسية ─────────────

  function buildBasic(ctx) {
    var ws = ctx.sheets.basic, T = ctx.T, N = T.NAMES, L = T.LAYOUT.basic, cfg = ctx.cfg;
    var org = ctx.project.org, a = ctx.analysis;
    T.setColumnWidths(ws, L.columns);
    titleBlock(ctx, ws, 'B', 'البيانات الأساسية', false);
    mergeStyled(ws, 'A2:B2', cfg.app.reportTitle, T.style.subtitle);

    var r = L.infoHeaderRow;
    mergeStyled(ws, 'A' + r + ':B' + r, 'بيانات الجهة', T.style.section);

    var actual = U.parseNumber(org.actualCount);
    var fields = [
      { label: 'اسم الجهة',          value: blankToNull(org.entity),     name: N.entity },
      { label: 'الإدارة',            value: blankToNull(org.department) },
      { label: 'القسم / الوحدة',     value: blankToNull(org.section) },
      { label: 'العدد الفعلي',       value: U.isValidNumber(actual) ? actual : null, name: N.actualCount, numFmt: T.NUM.integer,
        dv: numberValidation('whole', 'greaterThanOrEqual', 0, 'العدد الفعلي يجب أن يكون عددًا صحيحًا لا يقل عن صفر.') },
      { label: 'تاريخ الإعداد',      value: ctx.now, numFmt: T.NUM.date }   // تاريخ التصدير تلقائيًا
    ];

    fields.forEach(function (fd, i) {
      var row = r + 1 + i;
      var lc = ws.getCell('A' + row);
      lc.value = fd.label; T.style.label(lc);
      var vc = ws.getCell('B' + row);
      vc.value = fd.value;
      T.style.input(vc, fd.numFmt, fd.numFmt ? 'center' : 'right');
      if (fd.name) vc.addName(fd.name);
      if (fd.dv) vc.dataValidation = fd.dv;
      ws.getRow(row).height = 24;
    });

    var rr = r + fields.length + 1 + L.resultGapRows; // قسم النتائج بعد الحقول مباشرة مع سطر فاصل
    mergeStyled(ws, 'A' + rr + ':B' + rr, 'نتائج التحليل (محسوبة تلقائيًا)', T.style.section);
    var results = [
      { label: 'إجمالي ساعات العمل السنوية', v: f(N.totalHours, a.totalHours), fmt: T.NUM.hours },
      { label: 'الاحتياج المحسوب',           v: f(N.calcNeed, a.calcNeed), fmt: needFormat(cfg.needRounding.digits) },
      { label: 'الفجوة',                      v: f(N.gap, a.gap), fmt: gapFormat(cfg.needRounding.digits) },
      { label: 'الحالة',                      v: f(N.status, a.status), status: true }
    ];
    results.forEach(function (it, i) {
      var row = rr + 1 + i;
      var lc = ws.getCell('A' + row);
      lc.value = it.label; T.style.label(lc);
      var vc = ws.getCell('B' + row);
      vc.value = it.v;
      T.style.formula(vc, it.fmt, { bold: true });
      if (it.status) T.addStatusFormatting(ws, 'B' + row, cfg.status);
      ws.getRow(row).height = 24;
    });

    var lastRow = rr + results.length;
    T.applyPrint(ws, { paper: 'A4', landscape: false, printArea: 'A1:B' + lastRow, reportTitle: cfg.app.reportTitle, dateText: ctx.dateText });
  }

  // ───────────── ورقة تحليل عبء العمل ─────────────

  function buildWorkload(ctx) {
    var ws = ctx.sheets.workload, T = ctx.T, N = T.NAMES, L = T.LAYOUT.workload, cfg = ctx.cfg, a = ctx.analysis;
    T.setColumnWidths(ws, L.columns);
    titleBlock(ctx, ws, 'E', 'تحليل عبء العمل حسب التردد', true);

    var hr = L.headerRow;
    ['التردد', 'المرات السنوية', 'عدد المهام', 'إجمالي الساعات', 'نسبة الساعات'].forEach(function (h, i) {
      var c = ws.getRow(hr).getCell(i + 1); c.value = h; T.style.header(c);
    });
    ws.getRow(hr).height = 28;

    var first = hr + 1;
    a.byFrequency.forEach(function (bf, i) {
      var r = first + i;
      var stripe = i % 2 === 1 ? T.COLORS.stripe : T.COLORS.white;
      var row = ws.getRow(r);
      row.getCell(1).value = f('INDEX(' + N.freqNames + ',' + (i + 1) + ')', bf.label);
      row.getCell(2).value = f('INDEX(' + N.freqValues + ',' + (i + 1) + ')', bf.perYear);
      row.getCell(3).value = f('COUNTIF(' + N.taskFreq + ',A' + r + ')', bf.count);
      row.getCell(4).value = f('SUMIF(' + N.taskFreq + ',A' + r + ',' + N.taskHours + ')', bf.hours);
      row.getCell(5).value = f('IF(' + N.totalHours + '=0,0,D' + r + '/' + N.totalHours + ')', bf.share);
      T.style.formula(row.getCell(1), null, { fill: stripe, bold: true });
      T.style.formula(row.getCell(2), T.NUM.integer, { fill: stripe });
      T.style.formula(row.getCell(3), T.NUM.integer, { fill: stripe });
      T.style.formula(row.getCell(4), T.NUM.hours, { fill: stripe });
      T.style.formula(row.getCell(5), T.NUM.percent, { fill: stripe });
      row.height = 22;
    });

    var last = first + a.byFrequency.length - 1;
    var tr = last + 1;
    var tRow = ws.getRow(tr);
    tRow.getCell(1).value = 'الإجمالي';
    tRow.getCell(3).value = f('SUM(C' + first + ':C' + last + ')', a.taskCount);
    tRow.getCell(4).value = f('SUM(D' + first + ':D' + last + ')', a.totalHours);
    tRow.getCell(5).value = f('SUM(E' + first + ':E' + last + ')', a.totalHours > 0 ? 1 : 0);
    T.style.total(tRow.getCell(1));
    T.style.total(tRow.getCell(2));
    T.style.total(tRow.getCell(3), T.NUM.integer);
    T.style.total(tRow.getCell(4), T.NUM.hours);
    T.style.total(tRow.getCell(5), T.NUM.percent);
    tRow.height = 24;

    ws.autoFilter = 'A' + hr + ':E' + last;
    T.applyPrint(ws, { paper: 'A4', landscape: false, printArea: 'A1:E' + tr, titleRows: hr + ':' + hr, reportTitle: cfg.app.reportTitle, dateText: ctx.dateText });
  }

  // ───────────── ورقة الاحتياج والفجوة ─────────────

  function buildNeed(ctx) {
    var ws = ctx.sheets.need, T = ctx.T, N = T.NAMES, L = T.LAYOUT.need, cfg = ctx.cfg, a = ctx.analysis;
    var digits = cfg.needRounding.digits;
    T.setColumnWidths(ws, L.columns);
    titleBlock(ctx, ws, 'D', 'الاحتياج والفجوة', true);

    var hr = L.headerRow;
    ['الخطوة', 'البيان', 'القيمة', 'طريقة الحساب'].forEach(function (h, i) {
      var c = ws.getRow(hr).getCell(i + 1); c.value = h; T.style.header(c);
    });
    ws.getRow(hr).height = 28;

    var st = cfg.status;
    var steps = [
      { label: 'إجمالي ساعات العمل السنوية للمهام', v: f(N.totalHours, a.totalHours), fmt: T.NUM.hours,
        how: 'مجموع (المرات السنوية × التكرار × المدة بالدقائق ÷ ' + cfg.minutesPerHour + ') لجميع المهام' },
      { label: 'ساعات العمل الفعلية السنوية للموظف', v: f(N.annualHours, a.annualWorkHours), fmt: T.NUM.integer,
        how: 'من ورقة الإعدادات' },
      { label: 'الاحتياج قبل التقريب (للاطلاع)', name: N.exactNeed, fmt: T.NUM.decimal,
        v: f('IF(' + N.annualHours + '=0,0,' + N.totalHours + '/' + N.annualHours + ')', a.exactNeed),
        how: 'إجمالي الساعات ÷ ساعات العمل الفعلية السنوية' },
      { label: 'الاحتياج المحسوب', name: N.calcNeed, fmt: needFormat(digits), bold: true,
        v: f('ROUND(' + N.exactNeed + ',' + N.needDecimals + ')', a.calcNeed),
        how: 'تقريب لأقرب عدد صحيح: 0.5 فأعلى يُجبر للأعلى' },
      { label: 'العدد الفعلي', v: f(N.actualCount, a.actualCount), fmt: T.NUM.integer,
        how: 'من ورقة البيانات الأساسية' },
      { label: 'الفجوة', name: N.gap, fmt: gapFormat(digits), bold: true,
        v: f(N.calcNeed + '-' + N.actualCount, a.gap),
        how: 'الاحتياج المحسوب − العدد الفعلي' },
      { label: 'الحالة', name: N.status, bold: true, status: true,
        v: f('IF(' + N.gap + '>0,"' + st.deficit + '",IF(' + N.gap + '<0,"' + st.surplus + '","' + st.balanced + '"))', a.status),
        how: 'موجبة = ' + st.deficit + ' ، سالبة = ' + st.surplus + ' ، صفر = ' + st.balanced }
    ];

    steps.forEach(function (s, i) {
      var r = hr + 1 + i;
      var row = ws.getRow(r);
      row.getCell(1).value = i + 1;
      row.getCell(2).value = s.label;
      row.getCell(3).value = s.v;
      row.getCell(4).value = s.how;
      T.style.formula(row.getCell(1), null, { fill: T.COLORS.primaryLight, bold: true });
      T.style.label(row.getCell(2));
      T.style.formula(row.getCell(3), s.fmt, { bold: s.bold, size: s.bold ? 12 : 11 });
      T.style.formula(row.getCell(4), null, { align: 'right' });
      row.getCell(4).font = T.font({ size: 10, color: T.COLORS.muted });
      if (s.name) row.getCell(3).addName(s.name);
      if (s.status) T.addStatusFormatting(ws, 'C' + r, st);
      row.height = 26;
    });

    var lastRow = hr + steps.length;
    T.applyPrint(ws, { paper: 'A4', landscape: false, printArea: 'A1:D' + lastRow, reportTitle: cfg.app.reportTitle, dateText: ctx.dateText });
  }

  // ───────────── الملخص التنفيذي ─────────────

  function narrativeText(ctx) {
    var a = ctx.analysis;
    var txt = 'بناءً على تحليل ' + a.taskCount + ' مهمة بإجمالي ' + U.formatNumber(a.totalHours, 2) +
      ' ساعة عمل سنويًا، وبواقع ' + a.annualWorkHours + ' ساعة عمل فعلية سنوية للموظف، يبلغ الاحتياج المحسوب ' +
      a.calcNeed + ' موظف مقابل عدد فعلي ' + a.actualCount + '، وبذلك تكون الحالة: ' + a.status;
    return txt + (a.gap === 0 ? '.' : ' بمقدار ' + Math.abs(a.gap) + ' موظف.');
  }

  function buildSummary(ctx) {
    var ws = ctx.sheets.summary, T = ctx.T, N = T.NAMES, L = T.LAYOUT.summary, cfg = ctx.cfg, a = ctx.analysis;
    var digits = cfg.needRounding.digits, st = cfg.status;
    T.setColumnWidths(ws, L.columns);
    titleBlock(ctx, ws, 'H', 'الملخص التنفيذي', false);
    mergeStyled(ws, 'A2:H2', cfg.app.reportTitle, T.style.subtitle);

    var basic = function (addr) { return sheetRef(ctx, 'basic', addr); };
    var org = ctx.project.org;
    var deptText = [org.department, org.section].filter(function (x) { return String(x || '').trim(); }).join(' / ');
    var info = [
      { label: 'الجهة', v: f(N.entity, org.entity) },
      { label: 'الإدارة / القسم',
        v: f('IF(AND(' + basic('B6') + '="",' + basic('B7') + '=""),"—",' + basic('B6') + '&IF(AND(' + basic('B6') + '<>"",' + basic('B7') + '<>"")," / ","")&' + basic('B7') + ')', deptText || '—') }
    ];
    info.forEach(function (it, i) {
      var r = 4 + i;
      mergeStyled(ws, 'A' + r + ':B' + r, it.label, T.style.label);
      var vc = mergeStyled(ws, 'C' + r + ':H' + r, it.v);
      T.style.formula(vc, null, { align: 'right', bold: true });
      ws.getRow(r).height = 24;
    });

    /** صف مؤشرات على 8 أعمدة؛ كل مؤشر يأخذ عمودين افتراضيًا أو span أعمدة. */
    function kpiRow(labelRow, items) {
      var col = 0;
      items.forEach(function (it) {
        var span = it.span || 2;
        var c1 = String.fromCharCode(65 + col), c2 = String.fromCharCode(65 + col + span - 1);
        col += span;
        mergeStyled(ws, c1 + labelRow + ':' + c2 + labelRow, it.label, T.style.kpiLabel);
        var vc = mergeStyled(ws, c1 + (labelRow + 1) + ':' + c2 + (labelRow + 1), it.v);
        T.style.kpiValue(vc, it.fmt);
        if (it.status) T.addStatusFormatting(ws, c1 + (labelRow + 1), st);
      });
      ws.getRow(labelRow).height = 22;
      ws.getRow(labelRow + 1).height = 40;
    }

    kpiRow(8, [
      { label: 'عدد المهام', v: f(N.taskCount, a.taskCount), fmt: T.NUM.integer },
      { label: 'إجمالي ساعات العمل السنوية', v: f(N.totalHours, a.totalHours), fmt: T.NUM.hours },
      { label: 'ساعات العمل الفعلية السنوية', v: f(N.annualHours, a.annualWorkHours), fmt: T.NUM.integer },
      { label: 'الاحتياج المحسوب', v: f(N.calcNeed, a.calcNeed), fmt: needFormat(digits) }
    ]);
    kpiRow(11, [
      { label: 'العدد الفعلي', v: f(N.actualCount, a.actualCount), fmt: T.NUM.integer, span: 2 },
      { label: 'الفجوة', v: f(N.gap, a.gap), fmt: gapFormat(digits), span: 3 },
      { label: 'الحالة', v: f(N.status, a.status), status: true, span: 3 }
    ]);

    mergeStyled(ws, 'A14:H14', 'الخلاصة', T.style.section);
    var narrativeF =
      '"بناءً على تحليل "&' + N.taskCount + '&" مهمة بإجمالي "&FIXED(' + N.totalHours + ',2)&' +
      '" ساعة عمل سنويًا، وبواقع "&' + N.annualHours + '&" ساعة عمل فعلية سنوية للموظف، يبلغ الاحتياج المحسوب "&' +
      N.calcNeed + '&" موظف مقابل عدد فعلي "&' + N.actualCount + '&"، وبذلك تكون الحالة: "&' + N.status +
      '&IF(' + N.gap + '=0,"."," بمقدار "&ABS(' + N.gap + ')&" موظف.")';
    var p = mergeStyled(ws, 'A15:H17', f(narrativeF, narrativeText(ctx)), T.style.paragraph);
    ws.getRow(15).height = 24; ws.getRow(16).height = 24; ws.getRow(17).height = 24;
    p.protection = { locked: true };

    mergeStyled(ws, 'A19:H19', 'الاعتماد', T.style.section);
    var sign = [
      ['معد البيانات', null, 'التوقيع'],
      ['المعتمد', null, 'التوقيع'],
      ['التاريخ', null, '']
    ];
    sign.forEach(function (s, i) {
      var r = 20 + i;
      mergeStyled(ws, 'A' + r + ':B' + r, s[0], T.style.label);
      T.style.input(mergeStyled(ws, 'C' + r + ':D' + r, s[1]));
      if (s[2]) {
        mergeStyled(ws, 'E' + r + ':F' + r, s[2], T.style.label);
        T.style.input(mergeStyled(ws, 'G' + r + ':H' + r));
      }
      ws.getRow(r).height = 30;
    });

    T.applyPrint(ws, { paper: 'A4', landscape: false, printArea: 'A1:H22', reportTitle: cfg.app.reportTitle, dateText: ctx.dateText });
  }

  // ───────────── نطاق الطباعة الديناميكي ─────────────

  /**
   * ExcelJS لا يدعم نطاق طباعة مبنيًا على معادلة، لذا نمدد نموذج الأسماء المعرّفة لهذا المصنف فقط
   * (المكتبة مثبتة محليًا بإصدار 4.4.0). إذا لم تتوفر البنية المتوقعة يبقى نطاق الطباعة الثابت كما هو.
   */
  function applyDynamicPrintAreas(ctx) {
    var list = ctx.dynamicPrintAreas;
    if (!list.length) return;
    var dn = ctx.wb.definedNames;
    var desc = dn && Object.getOwnPropertyDescriptor(Object.getPrototypeOf(dn), 'model');
    if (!desc || typeof desc.get !== 'function') return;

    var order = ctx.cfg.excel.sheetOrder;
    var extra = list.map(function (p) {
      return { name: '_xlnm.Print_Area', localSheetId: order.indexOf(p.sheetKey), ranges: [p.formula] };
    });
    Object.defineProperty(dn, 'model', {
      configurable: true,
      get: function () { return desc.get.call(this).concat(extra); },
      set: function (v) { if (desc.set) desc.set.call(this, v); }
    });
    list.forEach(function (p) { ctx.sheets[p.sheetKey].pageSetup.printArea = undefined; });
  }

  // ───────────── التجميع ─────────────

  function protectionOptions() {
    return {
      selectLockedCells: true, selectUnlockedCells: true,
      formatColumns: true, formatRows: true, formatCells: false,
      insertRows: false, deleteRows: false,
      autoFilter: true, sort: true
    };
  }

  /**
   * يبني مصنف Excel كاملًا.
   * @returns {Promise<ExcelJS.Workbook>}
   */
  async function buildWorkbook(project, cfg, T, ExcelJSLib) {
    var ExcelJS = ExcelJSLib || root.ExcelJS;
    if (!ExcelJS) throw new Error('EXCELJS_MISSING');

    var wb = new ExcelJS.Workbook();
    var now = new Date();
    wb.creator = cfg.app.title;
    wb.title = cfg.app.reportTitle;
    wb.created = now;
    wb.modified = now;
    wb.calcProperties.fullCalcOnLoad = true;

    var ctx = {
      wb: wb, cfg: cfg, T: T, project: project,
      analysis: WL.calc.analyze(project, cfg),
      now: new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())),
      dateText: U.todayISO(),
      sheets: {},
      dynamicPrintAreas: []
    };

    var freeze = { tasks: T.LAYOUT.tasks.headerRow, workload: T.LAYOUT.workload.headerRow, need: T.LAYOUT.need.headerRow };
    cfg.excel.sheetOrder.forEach(function (key) {
      ctx.sheets[key] = wb.addWorksheet(cfg.excel.sheetNames[key], T.sheetOptions(freeze[key]));
    });

    // الترتيب مهم: الإعدادات أولًا ثم المهام، لأن باقي الأوراق تعتمد على أسمائهما
    buildSettings(ctx);
    buildTasks(ctx);
    buildBasic(ctx);
    buildWorkload(ctx);
    buildNeed(ctx);
    buildSummary(ctx);
    applyDynamicPrintAreas(ctx);

    if (cfg.excel.protection.enabled) {
      var pwd = cfg.excel.protection.password || '';
      for (var i = 0; i < cfg.excel.sheetOrder.length; i++) {
        await ctx.sheets[cfg.excel.sheetOrder[i]].protect(pwd, protectionOptions());
      }
    }
    return wb;
  }

  function buildFileName(project, cfg) {
    var entity = U.sanitizeFileName(project.org.entity) || 'جهة';
    return cfg.excel.fileNamePrefix + '_' + entity + '_' + U.todayISO() + '.xlsx';
  }

  async function exportToBuffer(project, cfg, T, ExcelJSLib) {
    var wb = await buildWorkbook(project, cfg, T, ExcelJSLib);
    return wb.xlsx.writeBuffer();
  }

  function downloadBuffer(buffer, fileName) {
    var blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  }

  WL.excel = {
    buildWorkbook: buildWorkbook,
    exportToBuffer: exportToBuffer,
    buildFileName: buildFileName,
    downloadBuffer: downloadBuffer
  };
})(typeof window !== 'undefined' ? window : globalThis);
