/**
 * sample-data.js — بيانات تجريبية واقعية لتجربة النظام واختباره.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});

  WL.sampleProject = function () {
    var M = WL.model;
    var p = M.createProject();
    p.org.entity = 'جامعة ........';   // اسم غير حقيقي: البيانات التجريبية لا تنتمي لجهة معروفة
    p.org.department = 'الإدارة العامة للموارد البشرية';
    p.org.section = 'قسم شؤون الموظفين';
    p.org.actualCount = '6';

    var ADMIN = 'مساعد إداري', HR = 'أخصائي موارد بشرية', PAYROLL = 'محاسب رواتب';
    [
      ['استقبال ومراجعة طلبات الموظفين الواردة عبر النظام', ADMIN, 'daily', 20, 10],
      ['إدخال وتحديث بيانات الموظفين في نظام الموارد البشرية', HR, 'daily', 15, 12],
      ['الرد على استفسارات الموظفين الهاتفية والبريدية', ADMIN, 'daily', 25, 6],
      ['إعداد خطابات التعريف والشهادات الوظيفية', ADMIN, 'daily', 8, 15],
      ['التدقيق على قرارات شؤون الموظفين قبل اعتمادها', HR, 'daily', 6, 20],
      ['حفظ وأرشفة الملفات الورقية والإلكترونية', ADMIN, 'daily', 10, 8],
      ['مراجعة سجلات الحضور والانصراف ومعالجة الملاحظات', PAYROLL, 'weekly', 10, 45],
      ['متابعة إجراءات التعيين والمباشرة', HR, 'weekly', 6, 90],
      ['معالجة طلبات الترقيات والعلاوات', HR, 'weekly', 5, 60],
      ['إعداد مسيرات العمل الإضافي والانتدابات', PAYROLL, 'semimonthly', 4, 120],
      ['إعداد تقرير الإجازات والغياب', PAYROLL, 'monthly', 3, 180],
      ['إعداد مسير الرواتب ومطابقته', PAYROLL, 'monthly', 2, 480],
      ['إعداد تقارير الأداء الربعية', HR, 'quarterly', 3, 600],
      ['تحديث الهيكل التنظيمي وبطاقات الوصف الوظيفي', HR, 'semiannual', 2, 900],
      ['إعداد الخطة السنوية للاحتياج التدريبي', HR, 'annual', 1, 2400],
      ['تقييم الأداء الوظيفي السنوي ومتابعة اعتماده', HR, 'annual', 1, 4800]
    ].forEach(function (t) {
      M.addTask(p, { title: t[0], positionTitle: t[1], frequencyKey: t[2], repetitions: String(t[3]), durationMinutes: String(t[4]) });
    });
    return p;
  };
})(typeof window !== 'undefined' ? window : globalThis);
