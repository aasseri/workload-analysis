/**
 * data-model.js — بنية بيانات المشروع وعمليات التعديل عليها.
 * القيم المشتقة (المرات السنوية، الساعات، النسب) لا تُخزّن هنا أبدًا؛ تُحسب في calculations.js.
 */
(function (root) {
  'use strict';
  var WL = (root.WL = root.WL || {});
  var U = WL.utils;

  var SCHEMA_VERSION = 1;

  var ORG_FIELDS = [
    'entity', 'department', 'section', 'jobTitle',
    'actualCount'
  ];

  function createProject() {
    return {
      schemaVersion: SCHEMA_VERSION,
      org: {
        entity: '',
        department: '',
        section: '',
        jobTitle: '',          // مخفي حاليًا في الواجهة؛ سياق اختياري للمساعد الذكي فقط، لا يُصدَّر إلى Excel
        actualCount: ''    // يُخزن كما أدخله المستخدم، ويُقرأ رقميًا عند التحقق/الحساب
        // تاريخ الإعداد لا يُخزن: يُكتب في ملف Excel بتاريخ التصدير تلقائيًا
      },
      tasks: [],
      employees: [] // محجوز للتوسع المستقبلي
    };
  }

  function createTask(partial) {
    var p = partial || {};
    return {
      id: U.uid(),
      title: p.title || '',
      frequencyKey: p.frequencyKey || '',
      repetitions: p.repetitions === undefined ? '' : p.repetitions,
      durationMinutes: p.durationMinutes === undefined ? '' : p.durationMinutes
    };
  }

  function indexOfTask(project, id) {
    for (var i = 0; i < project.tasks.length; i++) {
      if (project.tasks[i].id === id) return i;
    }
    return -1;
  }

  function addTask(project, partial) {
    var task = createTask(partial);
    project.tasks.push(task);
    return task;
  }

  function updateTask(project, id, changes) {
    var i = indexOfTask(project, id);
    if (i < 0) return null;
    Object.keys(changes).forEach(function (k) {
      if (k !== 'id') project.tasks[i][k] = changes[k];
    });
    return project.tasks[i];
  }

  function duplicateTask(project, id) {
    var i = indexOfTask(project, id);
    if (i < 0) return null;
    var copy = U.deepClone(project.tasks[i]);
    copy.id = U.uid();
    project.tasks.splice(i + 1, 0, copy);
    return copy;
  }

  function removeTask(project, id) {
    var i = indexOfTask(project, id);
    if (i < 0) return null;
    return project.tasks.splice(i, 1)[0];
  }

  /** ينقل المهمة إلى موضع جديد (فهرس يبدأ من 0). */
  function moveTaskTo(project, id, newIndex) {
    var i = indexOfTask(project, id);
    if (i < 0) return false;
    var target = Math.max(0, Math.min(project.tasks.length - 1, newIndex));
    if (target === i) return false;
    var task = project.tasks.splice(i, 1)[0];
    project.tasks.splice(target, 0, task);
    return true;
  }

  function moveTask(project, id, delta) {
    return moveTaskTo(project, id, indexOfTask(project, id) + delta);
  }

  function serialize(project) {
    return JSON.stringify(project, null, 2);
  }

  /** يقرأ مشروعًا محفوظًا ويكمل أي حقل ناقص بقيمه الافتراضية. يرمي خطأ إذا كان الملف غير صالح. */
  function deserialize(json) {
    var raw = typeof json === 'string' ? JSON.parse(json) : json;
    if (!raw || typeof raw !== 'object' || !raw.org || !Array.isArray(raw.tasks)) {
      throw new Error('INVALID_PROJECT');
    }
    var project = createProject();
    ORG_FIELDS.forEach(function (f) {
      if (raw.org[f] !== undefined && raw.org[f] !== null) project.org[f] = raw.org[f];
    });
    project.tasks = raw.tasks.map(function (t) {
      var task = createTask(t);
      if (t.id) task.id = String(t.id);
      return task;
    });
    project.employees = Array.isArray(raw.employees) ? raw.employees : [];
    return project;
  }

  WL.model = {
    SCHEMA_VERSION: SCHEMA_VERSION,
    ORG_FIELDS: ORG_FIELDS,
    createProject: createProject,
    createTask: createTask,
    indexOfTask: indexOfTask,
    addTask: addTask,
    updateTask: updateTask,
    duplicateTask: duplicateTask,
    removeTask: removeTask,
    moveTask: moveTask,
    moveTaskTo: moveTaskTo,
    serialize: serialize,
    deserialize: deserialize
  };
})(typeof window !== 'undefined' ? window : globalThis);
