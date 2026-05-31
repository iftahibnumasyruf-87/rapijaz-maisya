/* eslint-disable security/detect-object-injection */
/* eslint-disable i18next/no-literal-string */
import React, { useState, useEffect, useRef, useMemo, createContext, useContext } from 'react';
import { 
  Menu, X, Home, Users, BookOpen, Settings, LayoutTemplate, 
  Printer, CheckSquare, LogOut, Plus, Trash2, Edit2, Save,
  Download, Upload, Share2, AlertCircle, CheckCircle, GripHorizontal,
  Type, User, CreditCard, Image as ImageIcon, Ruler, Type as TypeIcon, FileText,
  Columns, FileSignature, TrendingUp, UserX, Clock, Activity, ChevronDown,
  ZoomIn, ZoomOut, Maximize, Minimize, ChevronUp, Lock, Database, Copy, Undo, Redo, Eye, EyeOff,
  AlignLeft, AlignCenter, AlignRight, AlignStartVertical, AlignCenterVertical, AlignEndVertical, BarChart2, AlignJustify, Layers, Calendar,
  Minus, Square
} from 'lucide-react';
import { createClient } from '@supabase/supabase-js';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LineChart, Line, AreaChart, Area } from 'recharts';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';

// ==========================================
// 1. SUPABASE SETUP (Koneksi Database Anda)
// ==========================================
const supabaseUrl = 'https://ikoqsyrvspfjyyjujfhc.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imlrb3FzeXJ2c3Bmanl5anVqZmhjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzczNjMyMzMsImV4cCI6MjA5MjkzOTIzM30.Q0IGVZFJr9Msaq-4pNgzilvH5Bu4zHoAXdrZFgmK45E';
const supabase = createClient(supabaseUrl, supabaseKey);

// ==========================================
// 1.5 UTILITY FUNCTIONS - KONVERSI ANGKA ARAB
// ==========================================
const convertArabicToLatin = (value) => {
  if (typeof value !== 'string') return value;
  
  return value
    .replace(/٠/g, '0')
    .replace(/١/g, '1')
    .replace(/٢/g, '2')
    .replace(/٣/g, '3')
    .replace(/٤/g, '4')
    .replace(/٥/g, '5')
    .replace(/٦/g, '6')
    .replace(/٧/g, '7')
    .replace(/٨/g, '8')
    .replace(/٩/g, '9');
};

// ==========================================
// 2. CONTEXT & STATE MANAGEMENT
// ==========================================
const AppContext = createContext();

const normalizeValue = (value) => {
  if (value === undefined || value === null) return '';
  return String(value).trim().toLowerCase();
};

const normalizeLookupValue = (value) => {
  if (value === undefined || value === null) return '';
  return String(value).trim().toLowerCase().replace(/[^a-z0-9\u0600-\u06FF]+/g, '');
};

const getClassNameFromValue = (classes, classValue) => {
  if (!classValue) return '';
  const normalizedInput = normalizeLookupValue(classValue);
  const found = classes.find(c => normalizeLookupValue(c.name) === normalizedInput || normalizeLookupValue(c.id) === normalizedInput);
  return found?.name || classValue;
};

const getClassIdFromValue = (classes, classValue) => {
  if (!classValue) return '';
  const normalizedInput = normalizeLookupValue(classValue);
  const found = classes.find(c => normalizeLookupValue(c.id) === normalizedInput || normalizeLookupValue(c.name) === normalizedInput);
  return found?.id || classValue;
};

const buildStableGradeDocId = (classId, activeSetting) => {
  if (!classId || !activeSetting) return null;
  const year = (activeSetting.tahun || 'default').replace(/\//g, '-');
  const semester = activeSetting.semester || '1';
  return `${classId}_${year}_${semester}`;
};

const buildLegacyGradeDocId = (className, activeSetting) => {
  if (!className || !activeSetting) return null;
  const year = (activeSetting.tahun || 'default').replace(/\//g, '-');
  const semester = activeSetting.semester || '1';
  return `${className}_${year}_${semester}`;
};

const getGradeDocId = (selectedClass, classes, activeSetting, grades = []) => {
  if (!selectedClass || !activeSetting) return null;
  const classId = getClassIdFromValue(classes, selectedClass);
  const className = getClassNameFromValue(classes, selectedClass);
  const stableId = buildStableGradeDocId(classId, activeSetting);
  const legacyNameId = buildLegacyGradeDocId(className, activeSetting);
  const fallbackId = `${selectedClass}_${(activeSetting.tahun || 'default').replace(/\//g, '-')}_${activeSetting.semester || '1'}`;
  if (grades.some(g => g.id === stableId)) return stableId;
  if (grades.some(g => g.id === legacyNameId)) return legacyNameId;
  if (grades.some(g => g.id === fallbackId)) return fallbackId;
  return stableId;
};

const arraysAreEqual = (a, b) => {
  if (!Array.isArray(a) || !Array.isArray(b)) return false;
  if (a.length !== b.length) return false;
  return a.every((value, index) => value === b[index]);
};

const normalizeSubjectClassIds = (classes, kelas) => {
  const values = normalizeSubjectClasses(kelas);
  return Array.from(new Set(values.map(v => getClassIdFromValue(classes, v) || v).filter(Boolean)));
};

const mergeGradePayloads = (targetPayload, sourcePayload) => {
  const merged = { ...targetPayload };
  merged.data = { ...targetPayload.data };

  Object.entries(sourcePayload.data || {}).forEach(([studentId, studentGrades]) => {
    if (!merged.data[studentId]) {
      merged.data[studentId] = studentGrades;
      return;
    }
    const existing = merged.data[studentId];
    if (typeof existing !== 'object' || existing === null) {
      merged.data[studentId] = existing ?? studentGrades;
      return;
    }
    if (typeof studentGrades !== 'object' || studentGrades === null) return;
    merged.data[studentId] = { ...studentGrades, ...existing };
  });

  if (!merged.class) merged.class = sourcePayload.class;
  if (!merged.tahun) merged.tahun = sourcePayload.tahun;
  if (!merged.semester) merged.semester = sourcePayload.semester;
  return merged;
};

const upsertPayload = async (colName, id, payload) => {
  const { error } = await supabase.from(colName).upsert([{ id, payload }]);
  return !error;
};

const deleteDocumentById = async (colName, id) => {
  const { error } = await supabase.from(colName).delete().eq('id', id);
  return !error;
};

const migrateLegacyData = async (newData) => {
  const activeSetting = newData.settings.find(s => s.isActive);
  if (!activeSetting) return false;

  let didMigrate = false;

  const colsToMigrateTahun = ['classes', 'subjectCategories', 'masterSubjects'];
  for (const col of colsToMigrateTahun) {
    if (Array.isArray(newData[col])) {
      for (const item of newData[col]) {
        if (!item.tahun || !item.semester) {
          const payload = { ...item, tahun: activeSetting.tahun, semester: activeSetting.semester };
          delete payload.id;
          const success = await upsertPayload(col, item.id, payload);
          if (success) didMigrate = true;
        }
      }
    }
  }
  if (Array.isArray(newData.students)) {
    for (const student of newData.students) {
      const normalizedId = getClassIdFromValue(newData.classes, student.kelas);
      if (normalizedId && normalizedId !== student.kelas) {
        const payload = { ...student, kelas: normalizedId };
        delete payload.id;
        const success = await upsertPayload('students', student.id, payload);
        if (success) didMigrate = true;
      }
    }
  }

  if (Array.isArray(newData.subjects)) {
    for (const subject of newData.subjects) {
      const normalizedIds = normalizeSubjectClassIds(newData.classes, subject.kelas);
      const originalIds = Array.isArray(subject.kelas) ? subject.kelas : normalizeSubjectClasses(subject.kelas);
      if (!arraysAreEqual(normalizedIds, originalIds)) {
        const payload = { ...subject, kelas: normalizedIds };
        delete payload.id;
        const success = await upsertPayload('subjects', subject.id, payload);
        if (success) didMigrate = true;
      }
    }
  }

  if (Array.isArray(newData.grades)) {
    const year = (activeSetting.tahun || 'default').replace(/\//g, '-');
    const semester = activeSetting.semester || '1';
    for (const cls of newData.classes) {
      const stableId = buildStableGradeDocId(cls.id, activeSetting);
      const legacyId = buildLegacyGradeDocId(cls.name, activeSetting);
      if (!stableId || stableId === legacyId) continue;
      const legacyDoc = newData.grades.find(g => g.id === legacyId);
      if (!legacyDoc) continue;
      const existingDoc = newData.grades.find(g => g.id === stableId);
      const legacyPayload = { ...legacyDoc };
      delete legacyPayload.id;

      if (existingDoc) {
        const currentPayload = { ...existingDoc };
        delete currentPayload.id;
        const mergedPayload = mergeGradePayloads(currentPayload, legacyPayload);
        mergedPayload.class = cls.id;
        mergedPayload.tahun = activeSetting.tahun;
        mergedPayload.semester = activeSetting.semester;
        const success = await upsertPayload('grades', stableId, mergedPayload);
        if (success) {
          await deleteDocumentById('grades', legacyId);
          didMigrate = true;
        }
      } else {
        const payload = { ...legacyPayload, class: cls.id, tahun: activeSetting.tahun, semester: activeSetting.semester };
        const success = await upsertPayload('grades', stableId, payload);
        if (success) {
          await deleteDocumentById('grades', legacyId);
          didMigrate = true;
        }
      }
    }
  }

  return didMigrate;
};

const getStudentsInClass = (students, classes, selectedClass) => {
  const normalizedSelected = normalizeLookupValue(selectedClass);
  if (!normalizedSelected) return [];

  const className = getClassNameFromValue(classes, selectedClass);
  const normalizedClassName = normalizeLookupValue(className);

  return students.filter((s) => {
    const studentClass = normalizeLookupValue(s.kelas);
    const studentClassName = getClassNameFromValue(classes, s.kelas);
    const normalizedStudentClassName = normalizeLookupValue(studentClassName);

    return (
      studentClass === normalizedSelected ||
      studentClass === normalizedClassName ||
      normalizedStudentClassName === normalizedSelected ||
      normalizedStudentClassName === normalizedClassName
    );
  });
};

// Mendapatkan data santri berdasarkan tahun ajaran aktif (snapshot jika ada, fallback ke data global)
const getStudentsForYear = (studentSnapshots = [], activeSetting, allStudents = []) => {
  if (!activeSetting || !activeSetting.tahun) return allStudents;
  const snapshotId = `${(activeSetting.tahun || '').replace(/\//g, '-')}_${activeSetting.semester || '1'}`;
  const snapshot = studentSnapshots.find(s => s.id === snapshotId);
  if (snapshot && Array.isArray(snapshot.students) && snapshot.students.length > 0) {
    return snapshot.students;
  }
  return allStudents;
};

const normalizeSubjectClasses = (kelas) => {
  if (Array.isArray(kelas)) {
    return kelas.map(c => String(c).trim()).filter(Boolean);
  }
  if (typeof kelas === 'string') {
    return kelas.split(',').map(c => c.trim()).filter(Boolean);
  }
  return [];
};

const getSubjectClassLabel = (subject, classes = []) => {
  const kelasValues = normalizeSubjectClasses(subject.kelas);
  if (kelasValues.length === 0) return 'Semua';
  const resolved = kelasValues.map((value) => getClassNameFromValue(classes, value) || value);
  return resolved.join(', ');
};

const parseGradeValue = (value) => {
  if (value === undefined || value === null || String(value).trim() === '') return null;
  const normalized = String(value).replace(/,/g, '.').trim();
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
};

const computeRaportScore = (uts, uas) => {
  const parsedUts = parseGradeValue(uts);
  const parsedUas = parseGradeValue(uas);
  if (parsedUts === null && parsedUas === null) return '';
  const u = parsedUts ?? 0;
  const a = parsedUas ?? 0;
  return Number((u * 0.4 + a * 0.6).toFixed(2));
};

const isSubjectVisibleInClass = (subject, selectedClass, classes = []) => {
  if (!selectedClass) return true;
  const normalizedSelected = normalizeLookupValue(selectedClass);
  const className = getClassNameFromValue(classes, selectedClass);
  const normalizedClassName = normalizeLookupValue(className);
  const kelasValues = normalizeSubjectClasses(subject.kelas);
  if (kelasValues.length === 0) return true;
  return kelasValues.some(c => {
    const normalizedClassValue = normalizeLookupValue(c);
    const mappedName = getClassNameFromValue(classes, c);
    const normalizedMappedName = normalizeLookupValue(mappedName);
    return (
      normalizedClassValue === normalizedSelected ||
      normalizedClassValue === normalizedClassName ||
      normalizedMappedName === normalizedSelected ||
      normalizedMappedName === normalizedClassName
    );
  });
};

const isReligiousCategory = (cat) => {
    if (!cat) return false;
    const n = normalizeValue(cat);
    const keywords = ['syari', 'syaria', 'syariyy', 'syariyyah', 'syar\u0131', 'syariah', 'agama', 'keagamaan', 'relig', 'islam'];
    return keywords.some(k => n.includes(k));
};

const sortSubjectsByCategory = (subjects, subjectCategories = []) => {
    const categoryOrder = new Map(subjectCategories.map((cat, idx) => [normalizeValue(cat.name), idx]));

    return [...subjects].sort((a, b) => {
        const aCategory = normalizeValue(a.kategori || '');
        const bCategory = normalizeValue(b.kategori || '');

        const aIsRel = isReligiousCategory(aCategory);
        const bIsRel = isReligiousCategory(bCategory);
        if (aIsRel !== bIsRel) return aIsRel ? -1 : 1;

        const aCatOrder = categoryOrder.has(aCategory) ? categoryOrder.get(aCategory) : Number.MAX_SAFE_INTEGER;
        const bCatOrder = categoryOrder.has(bCategory) ? categoryOrder.get(bCategory) : Number.MAX_SAFE_INTEGER;
        if (aCatOrder !== bCatOrder) return aCatOrder - bCatOrder;
        if (aCategory !== bCategory) return aCategory.localeCompare(bCategory);
        return normalizeValue(a.nameId || a.id).localeCompare(normalizeValue(b.nameId || b.id));
    });
};

const filterSubjectsByClass = (subjects, selectedClass, classes = []) => {
  const filtered = subjects.filter(subject => isSubjectVisibleInClass(subject, selectedClass, classes));
  const deduped = Array.from(
    new Map(filtered.map(s => [normalizeValue(s.nameId) || normalizeValue(s.id), s])).values()
  );
  return deduped;
};

const AppProvider = ({ children }) => {
  const [currentUser, setCurrentUser] = useState(() => {
    const saved = localStorage.getItem('rapijaz_user');
    return saved ? JSON.parse(saved) : null;
  });
  const currentUserRef = useRef(null);
  
  // Kolom yang terisolasi per semester (bukan data master global)
  const SEMESTER_SPECIFIC_COLS = ['students', 'subjects', 'teachers', 'extracurriculars', 'presences', 'characterTraits', 'classes', 'subjectCategories', 'masterSubjects'];

  const [allData, setAllData] = useState({
    settings: [], users: [], subjectCategories: [], masterSubjects: [], subjects: [], classes: [], students: [], teachers: [], 
    grades: [], layouts: [], fonts: [], studentFields: [], presences: [],
    extracurriculars: [], characterTraits: [], logs: [], studentSnapshots: []
  });

  // Semester yang sedang aktif
  const activeSetting = useMemo(() => allData.settings.find(s => s.isActive) || null, [allData.settings]);

  // Data yang sudah difilter sesuai semester aktif untuk komponen
  const data = useMemo(() => {
    if (!activeSetting) return allData;
    const filtered = { ...allData };
    for (const col of SEMESTER_SPECIFIC_COLS) {
      filtered[col] = allData[col].filter(item => {
        // Agar penambahan tahun ajaran baru datanya selalu baru/kosong,
        // kita hapus fallback yang menampilkan legacy data di semua semester.
        return item.tahun === activeSetting.tahun && item.semester === activeSetting.semester;
      });
    }
    return filtered;
  }, [allData, activeSetting]);

  const [loading, setLoading] = useState(true);
  const [notification, setNotification] = useState(null);
  const [autoSaveStatus, setAutoSaveStatus] = useState('idle'); // 'idle' | 'saving' | 'saved'

  useEffect(() => {
    currentUserRef.current = currentUser;
    if (currentUser) {
      localStorage.setItem('rapijaz_user', JSON.stringify(currentUser));
    } else {
      localStorage.removeItem('rapijaz_user');
    }
  }, [currentUser]);

  const showNotification = (message, type = 'success') => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 3000);
  };

  const sortDataItems = (items) => {
    return [...items].sort((a, b) => {
      const valA = String(a.name || a.nama || a.username || '').trim();
      const valB = String(b.name || b.nama || b.username || '').trim();
      if (!valA && !valB) return 0;
      return valA.localeCompare(valB, undefined, { numeric: true, sensitivity: 'base' });
    });
  };

  const fetchData = async (skipMigration = false) => {
    // PHASE 1: Load ONLY essential data for fast initial load
    const essentialCollections = ['settings', 'users', 'classes', 'layouts'];
    const lazyCollections = ['subjectCategories', 'masterSubjects', 'subjects', 'students', 'grades', 'fonts', 'studentFields', 'presences', 'extracurriculars', 'characterTraits', 'logs', 'teachers', 'studentSnapshots'];
    
    let newData = { ...allData };

    // Load essential data synchronously
    for (const colName of essentialCollections) {
      const { data: items, error } = await supabase.from(colName).select('*');
      if (!error && items) {
        newData[colName] = sortDataItems(items.map(item => ({ ...item.payload, id: item.id })));
      } else {
        newData[colName] = [];
      }
    }

    // Initialize lazy collections as empty
    for (const colName of lazyCollections) {
      newData[colName] = newData[colName] || [];
    }

    if (newData.users.length === 0) {
      const defaultAdmin = { username: 'admin', password: '123', role: 'admin', name: 'Administrator' };
      const { error } = await supabase.from('users').upsert([{ id: 'admin_1', payload: defaultAdmin }]);
      if (!error) {
         newData.users = [{ id: 'admin_1', ...defaultAdmin }];
      } else {
         console.error("Gagal membuat admin default di Supabase:", error);
      }
    }

    if (newData.layouts.length === 0) {
      const defaultLayouts = [
        { id: 'raport', name: 'Raport', elements: [], pageSize: 'A4', guides: { v: [], h: [] } },
        { id: 'ijazah', name: 'Ijazah', elements: [], pageSize: 'A4', guides: { v: [], h: [] } }
      ];
      for (const layout of defaultLayouts) {
        const { error } = await supabase.from('layouts').upsert([{ id: layout.id, payload: layout }]);
        if (!error) {
          newData.layouts.push(layout);
        }
      }
    }

    // Pastikan bucket 'layout-images' di Supabase Storage sudah ada
    try {
      const { data: buckets } = await supabase.storage.listBuckets();
      const bucketExists = buckets?.some(b => b.name === 'layout-images');
      if (!bucketExists) {
        await supabase.storage.createBucket('layout-images', { public: true, allowedMimeTypes: ['image/*'], fileSizeLimit: 5242880 });
      }
    } catch (_) { /* storage mungkin dibatasi oleh RLS, tidak masalah */ }

    if (!skipMigration) {
      const migrated = await migrateLegacyData(newData);
      if (migrated) {
        return fetchData(true);
      }
    }

    setAllData(newData);
    setLoading(false);

    // PHASE 2: Load lazy collections in background after UI renders
    setTimeout(() => fetchLazyCollections(newData), 500);
  };

  const fetchLazyCollections = async (currentData) => {
    const lazyCollections = ['subjectCategories', 'masterSubjects', 'subjects', 'students', 'grades', 'fonts', 'studentFields', 'presences', 'extracurriculars', 'characterTraits', 'logs', 'teachers'];
    let newData = { ...currentData };

    for (const colName of lazyCollections) {
      if (newData[colName].length === 0) {
        const { data: items, error } = await supabase.from(colName).select('*');
        if (!error && items) {
          newData[colName] = sortDataItems(items.map(item => ({ ...item.payload, id: item.id })));
        }
      }
    }

    setAllData(newData);
  };


  useEffect(() => {
    fetchData();
  }, []);

  const addLog = async (message) => {
    try {
      const logId = Date.now().toString();
      await supabase.from('logs').upsert([{
        id: logId, 
        payload: { message, timestamp: Date.now(), user: currentUserRef.current?.name || 'Sistem' }
      }]);
    } catch(e) {}
  };

  const saveToDb = async (colName, docId, payload, silent = false, customLogMsg = null) => {
    try {
      const { id: _payloadId, ...cleanPayload } = payload;

      // Jika koleksi terisolasi per semester, inject tahun & semester secara otomatis
      if (SEMESTER_SPECIFIC_COLS.includes(colName) && activeSetting && !cleanPayload.tahun) {
        cleanPayload.tahun = activeSetting.tahun;
        cleanPayload.semester = activeSetting.semester;
      }

      if (colName === 'settings' && cleanPayload.isActive) {
        const activeSettings = allData.settings.filter(s => s.id !== docId && s.isActive);
        if (activeSettings.length > 0) {
          const promises = activeSettings.map(s => {
            const { id, ...restPayload } = s;
            return supabase.from('settings').upsert([{ id: id, payload: { ...restPayload, isActive: false } }]);
          });
          await Promise.all(promises);
        }
      }

      const { error } = await supabase.from(colName).upsert([{ id: docId, payload: cleanPayload }]);
      if (error) throw error;
      
      if(!silent) showNotification('Data berhasil disimpan!');
      if(colName !== 'logs' && !silent) {
        addLog(customLogMsg || `Menyimpan data di menu ${colName}`);
      }
      
      // Only refetch the specific collection that was modified
      const { data: items } = await supabase.from(colName).select('*');
      if (items) {
        setAllData(prev => ({
          ...prev,
          [colName]: sortDataItems(items.map(item => ({ ...item.payload, id: item.id })))
        }));
      }
    } catch (err) {
      console.error('saveToDb error:', err);
      const msg = err?.message || '';
      if (!silent) {
        if (msg.includes('too large') || msg.includes('exceeded') || msg.includes('413')) {
          showNotification('Gagal simpan: Data terlalu besar. Gunakan gambar lebih kecil (< 200KB).', 'error');
        } else {
          showNotification(`Gagal menyimpan data. ${msg}`, 'error');
        }
      }
    }
  };

  const deleteFromDb = async (colName, docId, silent = false, customLogMsg = null) => {
    try {
      const { error } = await supabase.from(colName).delete().eq('id', docId);
      if (error) throw error;
      
      if(!silent) showNotification('Data berhasil dihapus!');
      if(colName !== 'logs' && !silent) {
        addLog(customLogMsg || `Menghapus data di menu ${colName}`);
      }
      
      // Only refetch the specific collection that was modified
      const { data: items } = await supabase.from(colName).select('*');
      if (items) {
        setAllData(prev => ({
          ...prev,
          [colName]: sortDataItems(items.map(item => ({ ...item.payload, id: item.id })))
        }));
      }
    } catch (err) {
      if(!silent) showNotification('Gagal menghapus data.', 'error');
    }
  };


  return (
    <AppContext.Provider value={{ data, allData, activeSetting, SEMESTER_SPECIFIC_COLS, currentUser, setCurrentUser, saveToDb, deleteFromDb, showNotification, addLog, autoSaveStatus, setAutoSaveStatus }}>
      {loading ? <div className="flex h-screen items-center justify-center bg-gray-50"><div className="animate-spin rounded-full h-12 w-12 border-b-2 border-emerald-600"></div></div> : children}
      {notification && (
        <div className={`fixed bottom-4 right-4 px-6 py-3 rounded-lg shadow-lg text-white font-medium flex items-center gap-2 z-50 transition-all ${notification.type === 'error' ? 'bg-red-500' : 'bg-emerald-600'}`}>
          {notification.type === 'error' ? <AlertCircle size={20} /> : <CheckCircle size={20} />}
          {notification.message}
        </div>
      )}
    </AppContext.Provider>
  );
};

const Modal = ({ isOpen, onClose, title, children }) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-xl shadow-xl w-full max-w-lg max-h-[90vh] flex flex-col">
        <div className="flex justify-between items-center p-4 border-b border-gray-100 shrink-0">
          <h2 className="text-xl font-bold text-gray-800">{title}</h2>
          <button onClick={onClose} className="p-1 hover:bg-gray-100 rounded-full"><X size={20} /></button>
        </div>
        <div className="p-4 overflow-y-auto flex-1 custom-scrollbar">{children}</div>
      </div>
    </div>
  );
};

const CurrentTime = () => {
    const [time, setTime] = useState(new Date());

    useEffect(() => {
        const timerId = setInterval(() => setTime(new Date()), 1000);
        return () => clearInterval(timerId);
    }, []);

    const getHijriDate = (date) => {
        const d = new Date(date);
        return new Intl.DateTimeFormat('id-TN-u-ca-islamic', {
            day: 'numeric',
            month: 'long',
            year: 'numeric'
        }).format(d);
    };

    const days = ['Ahad', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    const dayName = days[time.getDay()];
    const dateMasehi = time.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
    const timeString = time.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });

    return (
        <div className="flex flex-col items-end text-sm">
            <div className="font-bold text-emerald-800 flex items-center gap-1">
                 <Clock size={14} /> {timeString} WIB
            </div>
            <div className="text-gray-600 text-[11px] font-medium flex items-center gap-1">
                {dayName}, {dateMasehi} / {getHijriDate(time)}
            </div>
        </div>
    );
};

const AutoSaveIndicator = () => {
    const { autoSaveStatus } = useContext(AppContext);
    if (autoSaveStatus === 'idle') return null;

    return (
        <div className="flex items-center gap-1.5 text-gray-400 transition-all duration-300">
            {autoSaveStatus === 'saving' ? (
                <>
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"
                        className="animate-spin text-blue-400" style={{ width: 18, height: 18, animationDuration: '1s' }}>
                        <polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/>
                        <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>
                    </svg>
                    <span className="text-xs text-blue-400 font-medium">Menyimpan...</span>
                </>
            ) : (
                <>
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"
                        className="text-emerald-500" style={{ width: 18, height: 18 }}>
                        <path d="M12 2a5 5 0 0 1 4.9 4H18a4 4 0 0 1 0 8H6a4 4 0 0 1 0-8h1.1A5 5 0 0 1 12 2z"/>
                        <polyline points="9 12 11 14 15 10"/>
                    </svg>
                    <span className="text-xs text-emerald-500 font-medium">Tersimpan</span>
                </>
            )}
        </div>
    );
};

const Login = () => {
  const { data, setCurrentUser, showNotification, addLog } = useContext(AppContext);
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');

  const handleLogin = async (e) => {
    e.preventDefault();
    const user = data.users.find(u => u.username === username && u.password === password);
    if (user) {
      setCurrentUser(user);
      showNotification(`Selamat datang, ${user.name}`);
      try {
        const logId = Date.now().toString();
        await supabase.from('logs').upsert([{
          id: logId, 
          payload: { message: `Login berhasil`, timestamp: Date.now(), user: user.name }
        }]);
      } catch(err) {
        console.error("Gagal menyimpan log login", err);
      }
    } else {
      showNotification('Username atau password salah', 'error');
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-emerald-800 to-emerald-600 p-4">
      <div className="bg-white p-8 rounded-2xl shadow-2xl w-full max-w-md">
        <div className="text-center mb-8">
          <div className="mx-auto w-36 h-36 flex items-center justify-center mb-2">
            <img src="https://i.ibb.co.com/DfZSFRsP/Chat-GPT-Image-3-Mei-2026-04-08-56.png" alt="Logo Ponpes" className="w-full h-full object-contain drop-shadow-md" />
          </div>
          <h1 className="text-3xl font-bold text-gray-800 tracking-tight">Rapijaz-Maisya</h1>
          <p className="text-gray-500 mt-2 text-sm">Aplikasi Raport dan Ijazah<br/>Ponpes Imam Syafi'i Brebes</p>
        </div>
        <form onSubmit={handleLogin} className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Username</label>
            <input type="text" required className="w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-emerald-500 outline-none" value={username} onChange={e => setUsername(e.target.value)} />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Password</label>
            <input type="password" required className="w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-emerald-500 outline-none" value={password} onChange={e => setPassword(e.target.value)} />
          </div>
          <button type="submit" className="w-full bg-emerald-600 text-white py-2 rounded-lg font-semibold hover:bg-emerald-700 transition">Masuk</button>
        </form>
      </div>
    </div>
  );
};

const HomeDashboard = () => {
    const { data, saveToDb, deleteFromDb, currentUser, showNotification } = useContext(AppContext);
    const activeSetting = data.settings.find(s => s.isActive) || {};

    const [editLogModal, setEditLogModal] = useState(false);
    const [activeLog, setActiveLog] = useState({});
    
    // State untuk sorting tabel Log di Dashboard
    const [logSort, setLogSort] = useState({ key: 'timestamp', direction: 'descending' });

    const stats = useMemo(() => {
        let missingGrades = [];
        let topStudent = { name: '-', avg: 0, kelas: '-' };
        let topStudentsList = [];
        let mostAbsent = { name: '-', total: 0, kelas: '-', detail: '' };
        let classAvgMap = {};
        let subjectAvgMap = {};

        if (!activeSetting.tahun) return { missingGrades, topStudent, top10Students: [], mostAbsent, classAverages: [], subjectAverages: [] };

        const subjectIds = data.subjects.map(s => s.id);

        data.classes.forEach(c => {
            const gradeDocId = getGradeDocId(c.id, data.classes, activeSetting, data.grades);
            const classGrades = data.grades.find(g => g.id === gradeDocId)?.data || {};
            const studentsInClass = getStudentsInClass(data.students, data.classes, c.id);
            const className = getClassNameFromValue(data.classes, c.id);
            
            let classTotalScore = 0;
            let classTotalCount = 0;

            if (studentsInClass.length > 0) {
                data.subjects.forEach(sub => {
                    let hasGrade = false;
                    let subjTotal = 0;
                    let subjCount = 0;
                    studentsInClass.forEach(st => {
                        if (classGrades[st.id] && classGrades[st.id][sub.id] !== undefined && classGrades[st.id][sub.id] !== '') {
                            hasGrade = true;
                            const val = Number(classGrades[st.id][sub.id]);
                            if (!isNaN(val)) {
                                subjTotal += val;
                                subjCount++;
                            }
                        }
                    });
                    if (!hasGrade) {
                        missingGrades.push({ subject: sub.nameId, guru: sub.guru || '-', kelas: className });
                    }
                    if (subjCount > 0) {
                        if (!subjectAvgMap[sub.nameId]) subjectAvgMap[sub.nameId] = { total: 0, count: 0 };
                        subjectAvgMap[sub.nameId].total += subjTotal;
                        subjectAvgMap[sub.nameId].count += subjCount;
                    }
                });

                studentsInClass.forEach(st => {
                    const stGrades = classGrades[st.id] || {};
                    
                    let totalGrade = 0; let countGrade = 0;
                    subjectIds.forEach(sid => {
                        if (stGrades[sid] && !isNaN(stGrades[sid]) && stGrades[sid] !== '') {
                            totalGrade += Number(stGrades[sid]);
                            countGrade++;
                        }
                    });
                    const avg = countGrade > 0 ? (totalGrade / countGrade) : 0;
                    
                    if (countGrade > 0) {
                        topStudentsList.push({ name: st.nama, avg: Number(avg.toFixed(2)), kelas: className });
                        classTotalScore += totalGrade;
                        classTotalCount += countGrade;
                    }

                    if (avg > topStudent.avg) {
                        topStudent = { name: st.nama, avg: avg.toFixed(2), kelas: className };
                    }

                    let totalAbs = 0;
                    let absDetail = [];
                    data.presences.forEach(p => {
                        const val = Number(stGrades[p.id]);
                        if (!isNaN(val) && val > 0) {
                            totalAbs += val;
                            absDetail.push(`${val} ${p.name}`);
                        }
                    });
                    if (totalAbs > mostAbsent.total) {
                        mostAbsent = { name: st.nama, total: totalAbs, kelas: className, detail: absDetail.join(', ') };
                    }
                });
                
                if (classTotalCount > 0) {
                    classAvgMap[className] = classTotalScore / classTotalCount;
                }
            }
        });
        
        topStudentsList.sort((a, b) => b.avg - a.avg);
        const top10Students = topStudentsList.slice(0, 5); // Ambil Top 5
        
        const classAverages = Object.keys(classAvgMap).map(k => ({ name: k, 'Rata-Rata': Number(classAvgMap[k].toFixed(2)) }));
        const subjectAverages = Object.keys(subjectAvgMap).map(k => ({ name: k, 'Rata-Rata': Number((subjectAvgMap[k].total / subjectAvgMap[k].count).toFixed(2)) }));
        
        // Sort for better chart view
        classAverages.sort((a, b) => a.name.localeCompare(b.name));
        subjectAverages.sort((a, b) => b['Rata-Rata'] - a['Rata-Rata']); // Highest to lowest

        return { missingGrades, topStudent, top10Students, mostAbsent, classAverages, subjectAverages };
    }, [data, activeSetting]);

    // Logika pengurutan tabel Logs
    const sortedLogs = useMemo(() => {
        let sortableLogs = [...data.logs];
        sortableLogs.sort((a, b) => {
            let aValue = a[logSort.key];
            let bValue = b[logSort.key];
            
            if (typeof aValue === 'string') aValue = aValue.toLowerCase();
            if (typeof bValue === 'string') bValue = bValue.toLowerCase();
            
            if (aValue < bValue) return logSort.direction === 'ascending' ? -1 : 1;
            if (aValue > bValue) return logSort.direction === 'ascending' ? 1 : -1;
            return 0;
        });
        return sortableLogs;
    }, [data.logs, logSort]);

    const requestLogSort = (key) => {
        let direction = 'ascending';
        if (logSort.key === key && logSort.direction === 'ascending') {
            direction = 'descending';
        }
        setLogSort({ key, direction });
    };

    const LogHeader = ({ label, sortKey, className = "" }) => {
        const isActive = logSort.key === sortKey;
        return (
            <th className={`p-3 cursor-pointer select-none hover:bg-gray-200 transition-colors ${className}`} onClick={() => requestLogSort(sortKey)}>
                <div className="flex items-center justify-between gap-1">
                    <span>{label}</span>
                    <span className={`text-[10px] ${isActive ? 'text-blue-600 font-bold' : 'text-gray-400'}`}>
                        {isActive ? (logSort.direction === 'ascending' ? '▲' : '▼') : '↕'}
                    </span>
                </div>
            </th>
        );
    };

    const handleClearLogs = () => {
        if(window.confirm("Apakah Anda yakin ingin menghapus semua riwayat aktivitas?")) {
            data.logs.forEach(l => deleteFromDb('logs', l.id, true));
            showNotification("Seluruh riwayat aktivitas berhasil dibersihkan.");
        }
    };

    const handleSaveLog = () => {
        saveToDb('logs', activeLog.id, activeLog, true);
        setEditLogModal(false);
    };

    return (
        <div className="flex flex-col gap-6 h-[85vh] overflow-y-auto custom-scrollbar pb-6">
            <div className="bg-emerald-50 border border-emerald-100 rounded-xl p-6 shadow-sm">
                <h2 className="text-2xl font-bold text-emerald-800 flex items-center gap-2">
                    <span>🙏</span> Assalamu'alaikum, {currentUser?.name}
                </h2>
                <p className="text-emerald-600 mt-1">Selamat datang di Pusat Informasi Rapijaz-Maisya.</p>
            </div>

            {!activeSetting.tahun && (
                <div className="bg-yellow-50 border border-yellow-200 text-yellow-800 p-4 rounded-xl flex items-center gap-3">
                    <AlertCircle />
                    <div>
                        <p className="font-bold">Tahun Ajaran Belum Diatur</p>
                        <p className="text-sm">Mohon atur Tahun Ajaran dan Semester aktif di menu Master Data agar fitur analitik berjalan.</p>
                    </div>
                </div>
            )}

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 flex flex-col max-h-[300px]">
                    <div className="flex items-center gap-3 mb-4">
                        <div className="bg-red-100 p-2 rounded-lg text-red-600"><AlertCircle size={20}/></div>
                        <h3 className="font-bold text-gray-800">Belum Diinput</h3>
                    </div>
                    <div className="flex-1 overflow-y-auto custom-scrollbar pr-2 space-y-2">
                        {stats.missingGrades.length > 0 ? stats.missingGrades.map((mg, i) => (
                            <div key={i} className="bg-red-50/50 border border-red-100 p-3 rounded-lg text-sm">
                                <p className="font-bold text-gray-800">{mg.subject}</p>
                                <div className="flex justify-between text-gray-500 text-xs mt-1">
                                    <span>Guru: {mg.guru}</span>
                                    <span className="font-bold text-red-500">{mg.kelas}</span>
                                </div>
                            </div>
                        )) : (
                            <div className="h-full flex items-center justify-center text-emerald-600 text-sm font-bold flex-col gap-2">
                                <CheckCircle size={32} className="opacity-50"/>
                                Semua nilai telah diinput!
                            </div>
                        )}
                    </div>
                </div>

                <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 flex flex-col">
                    <div className="flex items-center gap-3 mb-4">
                        <div className="bg-emerald-100 p-2 rounded-lg text-emerald-600"><TrendingUp size={20}/></div>
                        <h3 className="font-bold text-gray-800">Nilai Tertinggi</h3>
                    </div>
                    <div className="flex-1 flex flex-col justify-center items-center text-center p-4 bg-emerald-50/30 rounded-xl border border-emerald-50">
                        <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mb-3">
                            <User size={32} />
                        </div>
                        <p className="text-xl font-bold text-gray-800">{stats.topStudent.name}</p>
                        <p className="text-gray-500 text-sm mb-3">Kelas: {stats.topStudent.kelas}</p>
                        <div className="bg-emerald-600 text-white px-4 py-1 rounded-full text-sm font-bold shadow-sm">
                            Rata-rata: {stats.topStudent.avg}
                        </div>
                    </div>
                </div>

                <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 flex flex-col">
                    <div className="flex items-center gap-3 mb-4">
                        <div className="bg-orange-100 p-2 rounded-lg text-orange-600"><UserX size={20}/></div>
                        <h3 className="font-bold text-gray-800">Ketidakhadiran Terbanyak</h3>
                    </div>
                    <div className="flex-1 flex flex-col justify-center items-center text-center p-4 bg-orange-50/30 rounded-xl border border-orange-50">
                        <div className="w-16 h-16 bg-orange-100 text-orange-600 rounded-full flex items-center justify-center mb-3">
                            <Clock size={32} />
                        </div>
                        <p className="text-xl font-bold text-gray-800">{stats.mostAbsent.name}</p>
                        <p className="text-gray-500 text-sm mb-3">Kelas: {stats.mostAbsent.kelas}</p>
                        <div className="bg-orange-500 text-white px-4 py-1 rounded-full text-sm font-bold shadow-sm">
                            Total: {stats.mostAbsent.total} Hari
                        </div>
                        <p className="text-[10px] text-gray-400 mt-2">({stats.mostAbsent.detail || '-'})</p>
                    </div>
                </div>
            </div>
            
            {activeSetting.tahun && (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mt-2">
                    <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100">
                        <div className="flex items-center gap-3 mb-4 border-b pb-3">
                            <div className="bg-blue-100 p-2 rounded-lg text-blue-600"><BarChart2 size={20}/></div>
                            <h3 className="font-bold text-gray-800">Rata-Rata Nilai per Kelas</h3>
                        </div>
                        <div className="h-[250px] w-full">
                            {stats.classAverages.length > 0 ? (
                                <ResponsiveContainer width="100%" height="100%">
                                    <BarChart data={stats.classAverages} margin={{ top: 10, right: 10, left: 0, bottom: 20 }}>
                                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
                                        <XAxis dataKey="name" tick={{fontSize: 12}} tickMargin={10} axisLine={false} tickLine={false} />
                                        <YAxis domain={[0, 100]} tick={{fontSize: 12}} axisLine={false} tickLine={false} />
                                        <Tooltip cursor={{fill: '#f8fafc'}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                                        <Bar dataKey="Rata-Rata" fill="#10b981" radius={[4, 4, 0, 0]} maxBarSize={50} />
                                    </BarChart>
                                </ResponsiveContainer>
                            ) : (
                                <div className="h-full flex items-center justify-center text-gray-400">Belum ada data nilai</div>
                            )}
                        </div>
                    </div>
                    
                    <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 flex flex-col">
                        <div className="flex items-center gap-3 mb-4 border-b pb-3">
                            <div className="bg-yellow-100 p-2 rounded-lg text-yellow-600"><Users size={20}/></div>
                            <h3 className="font-bold text-gray-800">Top 5 Santri Berprestasi</h3>
                        </div>
                        <div className="flex-1 overflow-x-auto">
                            <table className="w-full text-left text-sm border-collapse">
                                <thead className="bg-gray-50 text-gray-500">
                                    <tr>
                                        <th className="p-3 rounded-l-lg w-10 text-center">#</th>
                                        <th className="p-3">Nama Santri</th>
                                        <th className="p-3">Kelas</th>
                                        <th className="p-3 rounded-r-lg text-center">Rata-Rata</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {stats.top10Students.length > 0 ? stats.top10Students.map((st, idx) => (
                                        <tr key={idx} className="border-b border-gray-50 hover:bg-emerald-50/30 transition">
                                            <td className="p-3 text-center font-bold text-emerald-600">{idx + 1}</td>
                                            <td className="p-3 font-semibold text-gray-700">{st.name}</td>
                                            <td className="p-3 text-gray-500">{st.kelas}</td>
                                            <td className="p-3 text-center font-bold text-gray-800">{st.avg}</td>
                                        </tr>
                                    )) : (
                                        <tr><td colSpan="4" className="text-center py-8 text-gray-400">Belum ada data nilai</td></tr>
                                    )}
                                </tbody>
                            </table>
                        </div>
                    </div>
                    
                    <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 lg:col-span-2">
                        <div className="flex items-center gap-3 mb-4 border-b pb-3">
                            <div className="bg-purple-100 p-2 rounded-lg text-purple-600"><Activity size={20}/></div>
                            <h3 className="font-bold text-gray-800">Rata-Rata Nilai per Mata Pelajaran (Seluruh Kelas)</h3>
                        </div>
                        <div className="h-[300px] w-full">
                            {stats.subjectAverages.length > 0 ? (
                                <ResponsiveContainer width="100%" height="100%">
                                    <AreaChart data={stats.subjectAverages} margin={{ top: 10, right: 10, left: 0, bottom: 60 }}>
                                        <defs>
                                            <linearGradient id="colorRata" x1="0" y1="0" x2="0" y2="1">
                                                <stop offset="5%" stopColor="#8b5cf6" stopOpacity={0.3}/>
                                                <stop offset="95%" stopColor="#8b5cf6" stopOpacity={0}/>
                                            </linearGradient>
                                        </defs>
                                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
                                        <XAxis dataKey="name" tick={{fontSize: 11}} tickMargin={10} axisLine={false} tickLine={false} angle={-45} textAnchor="end" />
                                        <YAxis domain={[0, 100]} tick={{fontSize: 12}} axisLine={false} tickLine={false} />
                                        <Tooltip cursor={{stroke: '#e2e8f0', strokeWidth: 2}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                                        <Area type="monotone" dataKey="Rata-Rata" stroke="#8b5cf6" strokeWidth={3} fillOpacity={1} fill="url(#colorRata)" activeDot={{r: 6, strokeWidth: 0, fill: '#8b5cf6'}} />
                                    </AreaChart>
                                </ResponsiveContainer>
                            ) : (
                                <div className="h-full flex items-center justify-center text-gray-400">Belum ada data nilai</div>
                            )}
                        </div>
                    </div>
                </div>
            )}

            {currentUser?.role === 'admin' && (
                <>
                    <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 flex-1 flex flex-col min-h-[300px]">
                        <div className="flex justify-between items-center mb-4 pb-4 border-b">
                            <h3 className="font-bold text-gray-800 text-lg flex items-center gap-2"><Activity size={20} className="text-blue-500"/> Riwayat Penggunaan Aplikasi</h3>
                            <button onClick={handleClearLogs} className="text-sm bg-red-50 text-red-600 hover:bg-red-100 px-4 py-2 rounded-lg font-bold transition">
                                Hapus Semua Log
                            </button>
                        </div>
                        <div className="overflow-y-auto custom-scrollbar pr-2 flex-1">
                            <table className="w-full text-left text-sm border-collapse">
                                <thead className="bg-gray-50 text-gray-500 sticky top-0 shadow-sm z-10">
                                    <tr>
                                        <LogHeader label="Waktu" sortKey="timestamp" className="rounded-l-lg w-48" />
                                        <LogHeader label="Pengguna" sortKey="user" />
                                        <LogHeader label="Aktivitas" sortKey="message" />
                                        <th className="p-3 rounded-r-lg text-center w-24">Aksi</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {sortedLogs.length > 0 ? sortedLogs.map(log => (
                                        <tr key={log.id} className="border-b border-gray-50 hover:bg-blue-50/30 transition">
                                            <td className="p-3 text-gray-500">{new Date(log.timestamp).toLocaleString('id-ID')}</td>
                                            <td className="p-3 font-semibold text-gray-700">{log.user}</td>
                                            <td className="p-3 text-gray-800">{log.message}</td>
                                            <td className="p-3 text-center flex justify-center gap-2">
                                                <button onClick={() => { setActiveLog(log); setEditLogModal(true); }} className="text-blue-500 hover:bg-blue-100 p-1.5 rounded"><Edit2 size={14}/></button>
                                                <button onClick={() => deleteFromDb('logs', log.id, true)} className="text-red-500 hover:bg-red-100 p-1.5 rounded"><Trash2 size={14}/></button>
                                            </td>
                                        </tr>
                                    )) : (
                                        <tr><td colSpan="4" className="text-center py-8 text-gray-400">Belum ada riwayat aktivitas yang tercatat.</td></tr>
                                    )}
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <Modal isOpen={editLogModal} onClose={() => setEditLogModal(false)} title="Edit Riwayat (Log)">
                        <div className="space-y-4">
                            <div>
                                <label className="block text-sm font-medium text-gray-700 mb-1">Aktivitas</label>
                                <textarea className="w-full p-2 border rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" value={activeLog.message || ''} onChange={e => setActiveLog({...activeLog, message: e.target.value})} rows="3" />
                            </div>
                            <div>
                                <label className="block text-sm font-medium text-gray-700 mb-1">Nama Pengguna (Opsional ubah)</label>
                                <input className="w-full p-2 border rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" value={activeLog.user || ''} onChange={e => setActiveLog({...activeLog, user: e.target.value})} />
                            </div>
                            <div className="flex justify-end pt-4">
                                <button onClick={handleSaveLog} className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-lg font-bold transition">Simpan Perubahan</button>
                            </div>
                        </div>
                    </Modal>
                </>
            )}
        </div>
    );
};

const BackupRestorePanel = () => {
    const { data, allData, activeSetting, SEMESTER_SPECIFIC_COLS, saveToDb, showNotification } = useContext(AppContext);
    const [isProcessing, setIsProcessing] = useState(false);
    const [isCopyModalOpen, setIsCopyModalOpen] = useState(false);
    const [sourceSemesterId, setSourceSemesterId] = useState('');
    const [selectedCopyTargets, setSelectedCopyTargets] = useState([]);
    const [progressText, setProgressText] = useState('');
    
    const snapshotId = activeSetting ? `${(activeSetting.tahun || '').replace(/\//g, '-')}_${activeSetting.semester || '1'}` : null;
    const currentSnapshot = data.studentSnapshots.find(s => s.id === snapshotId);
    
    const handleSaveSnapshot = async () => {
        if (!activeSetting || !activeSetting.tahun) {
            showNotification('Aktifkan tahun ajaran terlebih dahulu!', 'error');
            return;
        }
        if (!confirm(`Simpan/Update Snapshot Santri untuk Tahun ${activeSetting.tahun} Semester ${activeSetting.semester}?`)) return;
        
        setIsProcessing(true);
        setProgressText('Mengecek database...');
        try {
            const { error: checkErr } = await supabase.from('studentSnapshots').select('id').limit(1);
            if (checkErr && checkErr.code === 'PGRST205') {
                 alert('GAGAL: Tabel "studentSnapshots" belum dibuat di database Supabase Anda.\n\nSilakan buka dashboard Supabase, buat tabel baru bernama "studentSnapshots", lalu tambahkan kolom: "id" (text/varchar, jadikan Primary Key) dan "payload" (jsonb).');
                 setIsProcessing(false);
                 return;
            }

            setProgressText('Menyimpan Snapshot...');
            const payload = {
                tahun: activeSetting.tahun,
                semester: activeSetting.semester,
                students: data.students,
                createdAt: new Date().toISOString()
            };
            await saveToDb('studentSnapshots', snapshotId, payload, false, `Menyimpan Snapshot Santri ${activeSetting.tahun}`);
            showNotification('Snapshot berhasil disimpan!');
        } catch (e) {
            console.error(e);
            showNotification('Gagal menyimpan snapshot', 'error');
        }
        setIsProcessing(false);
        setProgressText('');
    };

    const handleBackupExcel = async () => {
        if (!activeSetting || !activeSetting.tahun) {
            showNotification('Aktifkan Tahun Ajaran terlebih dahulu untuk mem-backup data!', 'error');
            return;
        }

        setIsProcessing(true);
        setProgressText('Mempersiapkan data backup...');
        
        try {
            await new Promise(r => setTimeout(r, 100)); // Biarkan UI merender overlay

            const wb = XLSX.utils.book_new();
            const collections = ['settings', 'users', 'subjectCategories', 'masterSubjects', 'subjects', 'classes', 'students', 'teachers', 'grades', 'layouts', 'fonts', 'studentFields', 'presences', 'extracurriculars', 'characterTraits', 'studentSnapshots'];
            
            for (const col of collections) {
                setProgressText(`Menyiapkan data: ${col}...`);
                await new Promise(r => setTimeout(r, 10)); 

                if (data[col] && data[col].length > 0) {
                    let filteredData = data[col];
                    if (['settings', 'grades', 'studentSnapshots'].includes(col)) {
                        filteredData = data[col].filter(item => item.tahun === activeSetting.tahun);
                    }
                    
                    if (filteredData.length === 0) continue;

                    const flatData = filteredData.map(item => {
                        let newItem = {};
                        for (const key in item) {
                            if (key === 'kelas' && data.classes && (col === 'students' || col === 'subjects')) {
                                if (Array.isArray(item[key])) {
                                    newItem[key] = item[key].map(k => getClassNameFromValue(data.classes, k) || k).join(', ');
                                } else {
                                    newItem[key] = getClassNameFromValue(data.classes, item[key]) || item[key];
                                }
                            } else if (typeof item[key] === 'object' && item[key] !== null) {
                                newItem[key] = JSON.stringify(item[key]);
                            } else {
                                newItem[key] = item[key];
                            }
                        }
                        return newItem;
                    });
                    const ws = XLSX.utils.json_to_sheet(flatData);
                    XLSX.utils.book_append_sheet(wb, ws, col);
                }
            }

            setProgressText('Membuat file Excel...');
            await new Promise(r => setTimeout(r, 50));

            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([wbout], { type: 'application/octet-stream' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const tahunStr = (activeSetting.tahun || 'unknown').replace(/\//g, '-');
            const semesterStr = activeSetting.semester || 'Semester';
            a.download = `backup_rapijaz_${tahunStr}_${semesterStr}.xlsx`;
            a.click();
            URL.revokeObjectURL(url);

            showNotification('Proses Backup Selesai!', 'success');
            alert('Proses Backup Berhasil!\n\nFile Excel telah diunduh ke perangkat Anda.');
        } catch (e) {
            console.error(e);
            showNotification('Gagal membuat file Excel', 'error');
            alert('Proses Backup Gagal!\n\nPastikan data Anda tidak ada yang korup.');
        } finally {
            setIsProcessing(false);
            setProgressText('');
        }
    };

    const handleRestoreExcel = (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        
        setIsProcessing(true);
        setProgressText('Membaca file Excel...');
        
        const reader = new FileReader();
        reader.onload = async (event) => {
            try {
                const fileData = new Uint8Array(event.target.result);
                const workbook = XLSX.read(fileData, {type: 'array'});
                
                let importedData = {};
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const rawData = XLSX.utils.sheet_to_json(worksheet);
                    importedData[sheetName] = rawData;
                });
                
                workbook.SheetNames.forEach(sheetName => {
                    importedData[sheetName] = importedData[sheetName].map(item => {
                        let parsedItem = {};
                        for (const key in item) {
                            if (key === 'kelas' && (sheetName === 'students' || sheetName === 'subjects')) {
                                const classList = importedData.classes || data.classes || [];
                                if (typeof item[key] === 'string' && item[key].includes(',')) {
                                    parsedItem[key] = item[key].split(',').map(k => getClassIdFromValue(classList, k.trim()) || k.trim());
                                } else {
                                    parsedItem[key] = getClassIdFromValue(classList, item[key]) || item[key];
                                }
                            } else if (typeof item[key] === 'string' && (item[key].startsWith('{') || item[key].startsWith('['))) {
                                try { parsedItem[key] = JSON.parse(item[key]); }
                                catch (err) { parsedItem[key] = item[key]; }
                            } else {
                                parsedItem[key] = item[key];
                            }
                        }
                        return parsedItem;
                    });
                });
                
                if (!importedData.settings || importedData.settings.length === 0) {
                    throw new Error("Format file backup tidak valid atau tidak memiliki data pengaturan (settings).");
                }
                
                // Validasi Tahun Ajaran
                const currentTahunSet = new Set(data.settings.map(s => String(s.tahun).trim()));
                const importedTahunSet = new Set(importedData.settings.map(s => String(s.tahun).trim()));
                
                let existingTahunWithData = [];
                for (let tahun of importedTahunSet) {
                    if (currentTahunSet.has(tahun)) {
                        let hasData = false;
                        const collectionsToCheck = ['students', 'subjects', 'teachers', 'grades', 'extracurriculars', 'presences', 'characterTraits'];
                        for (const col of collectionsToCheck) {
                            if (allData[col] && allData[col].some(item => item.tahun === tahun)) {
                                hasData = true;
                                break;
                            }
                        }
                        if (hasData) {
                            existingTahunWithData.push(tahun);
                        }
                    }
                }

                if (existingTahunWithData.length > 0) {
                    alert(`Tahun Ajaran ${existingTahunWithData.join(', ')} sudah ada di sistem dan sudah memiliki data. Restore dibatalkan untuk mencegah tertimpanya data. Anda hanya bisa me-restore ke Tahun Ajaran yang masih kosong.`);
                    e.target.value = '';
                    setIsProcessing(false);
                    setProgressText('');
                    return;
                }

                if (!confirm('Anda yakin ingin merestore data ini? (Aman: Data akan ditambahkan ke tahun ajaran yang masih kosong)')) {
                    e.target.value = '';
                    setIsProcessing(false);
                    setProgressText('');
                    return;
                }
                
                setProgressText('Memulai proses restore ke database...');
                
                const collections = ['settings', 'users', 'subjectCategories', 'masterSubjects', 'subjects', 'classes', 'students', 'teachers', 'grades', 'layouts', 'fonts', 'studentFields', 'presences', 'extracurriculars', 'characterTraits', 'studentSnapshots'];
                
                let count = 0;
                for (const col of collections) {
                    if (Array.isArray(importedData[col])) {
                         setProgressText(`Me-restore data: ${col}...`);
                         for (const item of importedData[col]) {
                             const { id, ...payload } = item;
                             if (id) {
                                 // eslint-disable-next-line no-await-in-loop
                                 await saveToDb(col, String(id), payload, true);
                                 count++;
                             }
                         }
                    }
                }
                
                setProgressText('Restore selesai. Memuat ulang...');
                showNotification(`Restore selesai (${count} data dipulihkan)! Halaman akan dimuat ulang.`);
                alert(`Proses Restore Berhasil!\n\nSebanyak ${count} data telah berhasil dimasukkan. Sistem akan dimuat ulang.`);
                setTimeout(() => window.location.reload(), 1500);
            } catch (err) {
                console.error(err);
                showNotification('Gagal memulihkan backup: ' + err.message, 'error');
                alert(`Proses Restore Gagal!\n\nError: ${err.message}`);
                setIsProcessing(false);
                setProgressText('');
            }
            e.target.value = '';
        };
        reader.onerror = () => {
             alert('Gagal membaca file backup!');
             setIsProcessing(false);
             setProgressText('');
        }
        reader.readAsArrayBuffer(file);
    };

    // Ambil daftar semester yang tersedia (untuk fitur Salin)
    const otherSemesters = allData.settings.filter(s => !s.isActive && s.tahun);

    const COPY_TARGET_LABELS = {
        students: 'Santri',
        subjects: 'Mata Pelajaran',
        teachers: 'Guru',
        extracurriculars: 'Ekstrakurikuler',
        presences: 'Aspek Presensi',
        characterTraits: 'Aspek Karakter',
        classes: 'Daftar Kelas',
        subjectCategories: 'Kategori Pelajaran',
        masterSubjects: 'Daftar Pelajaran Utama',
    };

    const toggleCopyTarget = (col) => {
        setSelectedCopyTargets(prev =>
            prev.includes(col) ? prev.filter(c => c !== col) : [...prev, col]
        );
    };

    const handleOpenCopyModal = () => {
        if (!activeSetting) {
            showNotification('Aktifkan semester tujuan terlebih dahulu!', 'error');
            return;
        }
        setSourceSemesterId(otherSemesters[0]?.id || '');
        setSelectedCopyTargets([...SEMESTER_SPECIFIC_COLS]);
        setIsCopyModalOpen(true);
    };

    const handleCopyFromSemester = async () => {
        if (!sourceSemesterId || selectedCopyTargets.length === 0) {
            showNotification('Pilih semester sumber dan data yang ingin disalin!', 'error');
            return;
        }
        const sourceSetting = allData.settings.find(s => s.id === sourceSemesterId);
        if (!sourceSetting) return;

        if (!confirm(`Salin data dari "${sourceSetting.tahun} ${sourceSetting.semester}" ke semester aktif saat ini (${activeSetting.tahun} ${activeSetting.semester})? Data yang sudah ada di semester aktif TIDAK akan ditimpa.`)) return;

        setIsProcessing(true);
        setIsCopyModalOpen(false);
        let count = 0;
        try {
            for (const col of selectedCopyTargets) {
                const sourceItems = allData[col].filter(item =>
                    item.tahun === sourceSetting.tahun && item.semester === sourceSetting.semester
                );
                // Cek item yang sudah ada di semester tujuan agar tidak ditimpa
                const existingNames = new Set(
                    allData[col]
                        .filter(item => item.tahun === activeSetting.tahun && item.semester === activeSetting.semester)
                        .map(item => (item.nama || item.name || item.username || '').toLowerCase())
                );
                for (const item of sourceItems) {
                    const itemName = (item.nama || item.name || item.username || '').toLowerCase();
                    if (existingNames.has(itemName)) continue; // Lewati jika sudah ada
                    const { id: _oldId, tahun: _t, semester: _s, ...rest } = item;
                    const newId = `${col}_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
                    // eslint-disable-next-line no-await-in-loop
                    await saveToDb(col, newId, { ...rest, tahun: activeSetting.tahun, semester: activeSetting.semester }, true);
                    count++;
                }
            }
            showNotification(`Berhasil menyalin ${count} data ke semester aktif!`);
        } catch (err) {
            showNotification('Gagal menyalin data: ' + err.message, 'error');
        }
        setIsProcessing(false);
    };

    return (
        <div className="space-y-6 max-w-4xl p-2">
            {/* Modal Salin dari Semester Lain */}
            {isCopyModalOpen && (
                <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
                    <div className="bg-white rounded-xl shadow-xl w-full max-w-lg">
                        <div className="flex justify-between items-center p-4 border-b">
                            <h2 className="text-xl font-bold text-gray-800">Salin Data dari Semester Lain</h2>
                            <button onClick={() => setIsCopyModalOpen(false)} className="p-1 hover:bg-gray-100 rounded-full"><X size={20}/></button>
                        </div>
                        <div className="p-5 space-y-4">
                            <div>
                                <label className="block text-sm font-semibold text-gray-700 mb-1">Semester Sumber (Asal Data)</label>
                                <select
                                    className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm"
                                    value={sourceSemesterId}
                                    onChange={e => setSourceSemesterId(e.target.value)}
                                >
                                    {otherSemesters.map(s => (
                                        <option key={s.id} value={s.id}>{s.tahun} - {s.semester}</option>
                                    ))}
                                </select>
                                {otherSemesters.length === 0 && <p className="text-xs text-red-500 mt-1">Tidak ada semester lain yang tersedia.</p>}
                            </div>
                            <div>
                                <label className="block text-sm font-semibold text-gray-700 mb-2">Data yang Ingin Disalin</label>
                                <div className="grid grid-cols-2 gap-2">
                                    {SEMESTER_SPECIFIC_COLS.map(col => (
                                        <label key={col} className="flex items-center gap-2 text-sm cursor-pointer bg-gray-50 hover:bg-gray-100 px-3 py-2 rounded-lg">
                                            <input
                                                type="checkbox"
                                                checked={selectedCopyTargets.includes(col)}
                                                onChange={() => toggleCopyTarget(col)}
                                                className="w-4 h-4 text-emerald-600"
                                            />
                                            {COPY_TARGET_LABELS[col]}
                                        </label>
                                    ))}
                                </div>
                            </div>
                            <p className="text-xs text-gray-500 bg-amber-50 border border-amber-200 rounded-lg p-3">
                                ⚠️ Data yang namanya sudah ada di semester aktif TIDAK akan ditimpa. Hanya data baru yang akan disalin.
                            </p>
                        </div>
                        <div className="flex gap-3 p-4 border-t">
                            <button
                                onClick={handleCopyFromSemester}
                                disabled={isProcessing || otherSemesters.length === 0 || selectedCopyTargets.length === 0}
                                className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2.5 rounded-lg font-bold transition disabled:opacity-50"
                            >
                                {isProcessing ? 'Menyalin...' : 'Mulai Salin Data'}
                            </button>
                            <button onClick={() => setIsCopyModalOpen(false)} className="px-4 py-2.5 rounded-lg border font-medium hover:bg-gray-50">Batal</button>
                        </div>
                    </div>
                </div>
            )}

            <div className="bg-blue-50 border border-blue-200 p-6 rounded-xl">
                <h4 className="font-bold text-blue-900 text-lg mb-2 flex items-center gap-2"><Lock size={20}/> Kunci Data Santri (Snapshot)</h4>
                <p className="text-sm text-blue-800 mb-4">
                    Fitur ini digunakan untuk <b>mengunci data santri</b> pada tahun ajaran yang sedang aktif. 
                    Saat tahun ajaran berganti, santri yang sudah naik kelas atau lulus tidak akan merubah tampilan raport lama.
                </p>
                <div className="flex items-center justify-between bg-white p-4 rounded-lg shadow-sm border border-blue-100">
                    <div>
                        <p className="font-semibold text-gray-800">Status Snapshot: {activeSetting ? `${activeSetting.tahun} Sem ${activeSetting.semester}` : 'Tidak ada tahun aktif'}</p>
                        <p className="text-sm text-gray-500">
                            {currentSnapshot ? `✅ Terkunci dengan ${currentSnapshot.students?.length || 0} santri (Diperbarui: ${new Date(currentSnapshot.createdAt).toLocaleString()})` : '⚠️ Belum ada snapshot untuk tahun ini.'}
                        </p>
                    </div>
                    <button 
                        onClick={handleSaveSnapshot} 
                        disabled={isProcessing}
                        className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2.5 rounded-lg font-bold transition shadow-sm"
                    >
                        {isProcessing ? 'Menyimpan...' : 'Simpan Snapshot'}
                    </button>
                </div>
            </div>

            {/* Salin Data dari Semester Lain */}
            <div className="bg-violet-50 border border-violet-200 p-6 rounded-xl">
                <h4 className="font-bold text-violet-900 text-lg mb-2 flex items-center gap-2"><Share2 size={20}/> Salin Data dari Semester Lain</h4>
                <p className="text-sm text-violet-800 mb-4">
                    Gunakan fitur ini untuk menyalin data (Santri, Mata Pelajaran, Guru, dll) dari semester lain ke semester yang sedang aktif. Data yang sudah ada di semester aktif tidak akan ditimpa.
                </p>
                <button
                    onClick={handleOpenCopyModal}
                    disabled={isProcessing || otherSemesters.length === 0}
                    className="bg-violet-600 hover:bg-violet-700 text-white px-4 py-3 rounded-lg font-bold flex items-center gap-2 transition disabled:opacity-50"
                >
                    <Share2 size={20}/> {otherSemesters.length === 0 ? 'Belum ada semester lain' : 'Pilih & Salin Data...'}
                </button>
            </div>

            <div className="bg-emerald-50 border border-emerald-200 p-6 rounded-xl">
                <h4 className="font-bold text-emerald-900 text-lg mb-2 flex items-center gap-2"><Database size={20}/> Backup & Restore Seluruh Data</h4>
                <p className="text-sm text-emerald-800 mb-4">
                    Gunakan fitur ini untuk mencadangkan (backup) seluruh data sistem dalam format Excel (.xlsx), mengeditnya secara manual, lalu memulihkannya (restore).
                </p>
                <div className="flex gap-4">
                    <button onClick={handleBackupExcel} className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-3 rounded-lg font-bold flex items-center justify-center gap-2 transition shadow-sm">
                        <Download size={20}/> Download Backup (.xlsx)
                    </button>
                    
                    <div className="flex-1 relative cursor-pointer">
                         <input 
                            type="file" 
                            accept=".xlsx, .xls" 
                            onChange={handleRestoreExcel} 
                            disabled={isProcessing}
                            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                        />
                        <div className={`bg-gray-800 hover:bg-gray-900 text-white px-4 py-3 rounded-lg font-bold flex items-center justify-center gap-2 transition shadow-sm ${isProcessing ? 'opacity-50' : ''}`}>
                            <Upload size={20}/> {isProcessing ? 'Memproses...' : 'Restore Data (.xlsx)'}
                        </div>
                    </div>
                </div>
            </div>

            {/* OVERLAY LOADING */}
            {isProcessing && (
                <div className="fixed inset-0 bg-gray-900/80 backdrop-blur-sm z-[9999] flex flex-col items-center justify-center pointer-events-auto">
                    <div className="bg-white p-8 rounded-2xl shadow-2xl flex flex-col items-center max-w-sm w-full mx-4 text-center">
                        <div className="animate-spin rounded-full h-16 w-16 border-4 border-emerald-100 border-b-emerald-600 mb-6"></div>
                        <h3 className="text-xl font-bold text-gray-800 mb-2">Sistem Terkunci sementara</h3>
                        <p className="text-emerald-600 font-medium">{progressText || 'Memproses permintaan Anda...'}</p>
                        <p className="text-sm text-gray-400 mt-4">Mohon jangan tutup halaman ini selama proses berlangsung.</p>
                    </div>
                </div>
            )}
        </div>
    );
};

const MasterData = ({ activeTab }) => {
  const { data, allData, saveToDb, deleteFromDb, showNotification } = useContext(AppContext);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingItem, setEditingItem] = useState(null);
  const [formData, setFormData] = useState({});
  const [isAutoSaving, setIsAutoSaving] = useState(false);
    const [autoSaveEnabled, setAutoSaveEnabled] = useState(false);
    const [hideInactive, setHideInactive] = useState(false);
    const [isBulkProcessing, setIsBulkProcessing] = useState(false);
    const [bulkProgressText, setBulkProgressText] = useState('');
    const [bulkProgressCurrent, setBulkProgressCurrent] = useState(0);
    const [bulkProgressTotal, setBulkProgressTotal] = useState(0);

  // Default pengurutan tabel berdasarkan menu (alfabetis secara bawaan)
  const getDefaultSortKey = (tab) => {
      switch(tab) {
          case 'settings': return 'tahun';
          case 'teachers': return 'nama';
          case 'subjectCategories': return 'name';
          case 'masterSubjects': return 'nameId';
          case 'subjects': return 'kelas';
          case 'presences': return 'name';
          case 'characterTraits': return 'name';
          case 'extracurriculars': return 'name';
          case 'fonts': return 'name';
          case 'studentFields': return 'name';
          case 'students': return 'nama';
          case 'classes': return 'name';
          case 'users': return 'name';
          default: return 'id';
      }
  };

  const [sortConfig, setSortConfig] = useState({ key: getDefaultSortKey(activeTab), direction: 'ascending' });

  // Reset tabel sorting saat berpindah menu (tab)
  useEffect(() => {
      setSortConfig({ key: getDefaultSortKey(activeTab), direction: 'ascending' });
  }, [activeTab]);

  const requestSort = (key) => {
      let direction = 'ascending';
      if (sortConfig.key === key && sortConfig.direction === 'ascending') {
          direction = 'descending';
      }
      setSortConfig({ key, direction });
  };

  const sortedData = useMemo(() => {
      let sortableItems = [...(data[activeTab] || [])];
      if (sortConfig.key) {
          sortableItems.sort((a, b) => {
              let aValue = a[sortConfig.key];
              let bValue = b[sortConfig.key];
              
              if (aValue === undefined || aValue === null) aValue = '';
              if (bValue === undefined || bValue === null) bValue = '';
              
              if (activeTab === 'subjects' && sortConfig.key === 'kelas') {
                  aValue = getSubjectClassLabel(a, allData?.classes || data.classes);
                  bValue = getSubjectClassLabel(b, allData?.classes || data.classes);
              }

              if (typeof aValue === 'string' && typeof bValue === 'string') {
                  let comparison = aValue.localeCompare(bValue, undefined, { numeric: true, sensitivity: 'base' });
                  
                  // Secondary sort untuk halaman Plotting Pelajaran (berdasarkan Kategori) jika kelasnya sama
                  if (comparison === 0 && activeTab === 'subjects' && sortConfig.key === 'kelas') {
                      const catA = a.kategori || '';
                      const catB = b.kategori || '';
                      comparison = catA.localeCompare(catB, undefined, { numeric: true, sensitivity: 'base' });
                  }

                  return sortConfig.direction === 'ascending' ? comparison : -comparison;
              }
              
              if (aValue < bValue) return sortConfig.direction === 'ascending' ? -1 : 1;
              if (aValue > bValue) return sortConfig.direction === 'ascending' ? 1 : -1;
              return 0;
          });
      }
      return sortableItems;
  }, [data, activeTab, sortConfig]);

  const groupedSubjects = useMemo(() => {
      if (activeTab !== 'subjects') return sortedData;
      const groups = [];
      let currentKelas = null;
      let classIndex = 0;
      sortedData.forEach(sub => {
          const kelasLabel = getSubjectClassLabel(sub, allData?.classes || data.classes);
          if (kelasLabel !== currentKelas) {
              currentKelas = kelasLabel;
              classIndex = 0;
              groups.push({ type: 'group', kelas: kelasLabel });
          }
          classIndex += 1;
          groups.push({ type: 'item', subject: sub, number: classIndex });
      });
      return groups;
  }, [sortedData, activeTab]);

  const SortableHeader = ({ label, sortKey, className = "" }) => {
      const isActive = sortConfig.key === sortKey;
      return (
          <th className={`p-3 border-b cursor-pointer select-none hover:bg-gray-200 transition-colors ${className}`} onClick={() => requestSort(sortKey)}>
              <div className="flex items-center justify-between gap-2">
                  <span>{label}</span>
                  <span className={`text-[10px] ${isActive ? 'text-emerald-600 font-bold' : 'text-gray-400'}`}>
                      {isActive ? (sortConfig.direction === 'ascending' ? '▲' : '▼') : '↕'}
                  </span>
              </div>
          </th>
      );
  };

  const handleOpenModal = (item = null) => {
    const newItem = item ? { ...item } : { id: Date.now().toString() };
    if (item && activeTab === 'subjects') {
      const kelasValues = normalizeSubjectClasses(item.kelas);
      newItem.kelas = kelasValues.map(k => getClassIdFromValue(data.classes, k));
    }
    if (item && activeTab === 'students') {
      newItem.kelas = getClassIdFromValue(data.classes, item.kelas);
    }
    setEditingItem(newItem);
    setFormData(newItem);
    setIsModalOpen(true);
  };

  const handleSave = () => {
    let payload = { ...formData };
    
    if (activeTab === 'settings' && !editingItem.tahun) {
      // Saat membuat tahun ajaran baru, otomatis buat 2 semester (Ganjil & Genap)
      const id1 = formData.id || Date.now().toString();
      const id2 = (parseInt(id1) + 1).toString();
      
      const ganjilPayload = {
        ...payload,
        semester: 'Ganjil',
        semester_arab: 'الفصل الدراسي الأول',
        isActive: payload.isActive || false
      };
      
      const genapPayload = {
        ...payload,
        semester: 'Genap',
        semester_arab: 'الفصل الدراسي الثاني',
        isActive: false
      };

      saveToDb(activeTab, id1, ganjilPayload, true);
      saveToDb(activeTab, id2, genapPayload, false, `Menambah Tahun Ajaran Baru (${payload.tahun})`);
      setIsModalOpen(false);
      return;
    }

    if (activeTab === 'users' && !payload.role) payload.role = 'user';
    saveToDb(activeTab, formData.id, payload, false, `Menyimpan data di Master Data (${activeTab})`);
    setIsModalOpen(false);
  };

    useEffect(() => {
        if (!isModalOpen) return;
        if (!autoSaveEnabled) return; // skip autosave when disabled
        const timer = setTimeout(() => {
                // CEK KOTOR (DIRTY CHECK): Cek apakah data benar-benar berubah dari data acuan (editingItem)
                const hasChanged = JSON.stringify(formData) !== JSON.stringify(editingItem);
        
                // HANYA autosave jika ada minimal 1 field terisi DAN terjadi perubahan nyata
                if (Object.keys(formData).length > 1 && hasChanged) {
                        setIsAutoSaving(true);
                        let payload = { ...formData };
                        if (activeTab === 'users' && !payload.role) payload.role = 'user';
            
                        saveToDb(activeTab, formData.id, payload, true);
                        setEditingItem(payload); // Set data saat ini sebagai acuan baru agar tidak disave ulang terus-menerus
            
                        setTimeout(() => setIsAutoSaving(false), 800);
                }
        }, 5000); 
        return () => clearTimeout(timer);
    }, [formData, isModalOpen, activeTab, saveToDb, editingItem, autoSaveEnabled]);

  const translateToArabic = async (text) => {
      if (!text) return '';
      try {
          const res = await fetch(`https://translate.googleapis.com/translate_a/single?client=gtx&sl=id&tl=ar&dt=t&q=${encodeURIComponent(text)}`);
          const json = await res.json();
          return json[0][0][0] || '';
      } catch (e) {
          console.error('Terjadi kesalahan translasi:', e);
          return '';
      }
  };

  const convertLatinDigitsToArabic = (value) => {
      if (typeof value !== 'string') return value;
      const map = { '0': '٠', '1': '١', '2': '٢', '3': '٣', '4': '٤', '5': '٥', '6': '٦', '7': '٧', '8': '٨', '9': '٩' };
      return value.replace(/[0-9]/g, digit => map[digit] || digit);
  };

    const generateExcelTemplate = (type) => {
        const headers = ['NIS', 'Nama Lengkap', 'Nama Arab', 'Kelas'];
        try {
            (data.studentFields || []).forEach(f => {
                if (f && f.name) headers.push(f.name);
            });
        } catch (e) {}

        const ws = XLSX.utils.aoa_to_sheet([headers]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'TemplateSantri');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'template_data_santri.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    };

    const handleImportExcel = async (e, type) => {
        const file = e.target.files[0];
        if (!file) return;
        e.target.value = '';
        setIsBulkProcessing(true);
        setBulkProgressText('Membaca file Excel...');
        setBulkProgressCurrent(0);
        setBulkProgressTotal(0);
        try {
            await new Promise(r => setTimeout(r, 80));
            const dataBuf = await file.arrayBuffer();
            const workbook = XLSX.read(dataBuf, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
            const total = rows.length;
            setBulkProgressTotal(total);
            let count = 0;
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const item = {};
                Object.keys(row).forEach(k => {
                    const key = k.toString().trim().toLowerCase();
                    const val = row[k];
                    if (key === 'nis') item.nis = val;
                    else if (key.includes('nama') && key.includes('arab')) item.nama_arab = val;
                    else if (key.includes('nama')) item.nama = val;
                    else if (key === 'kelas') item.kelas = val;
                    else item[key] = val;
                });
                item.id = Date.now().toString() + i;
                setBulkProgressText(`Mengimpor santri: ${item.nama || `Data ke-${i+1}`}`);
                setBulkProgressCurrent(i + 1);
                // eslint-disable-next-line no-await-in-loop
                await saveToDb(type, item.id, item, true);
                count++;
            }
            setBulkProgressText('Selesai!');
            showNotification(`${count} data berhasil diimpor dari Excel!`);
        } catch (err) {
            console.error(err);
            showNotification('Gagal memproses file Excel. Pastikan format benar.', 'error');
        } finally {
            setIsBulkProcessing(false);
            setBulkProgressText('');
            setBulkProgressCurrent(0);
            setBulkProgressTotal(0);
        }
    };

  const handleDeleteAllStudents = async () => {
    const total = data.students?.length || 0;
    if (total === 0) {
      showNotification('Tidak ada data santri untuk dihapus.', 'error');
      return;
    }
    const firstConfirm = confirm(`⚠️ PERINGATAN!\n\nAnda akan menghapus SEMUA ${total} data santri secara permanen.\n\nApakah Anda yakin ingin melanjutkan?`);
    if (!firstConfirm) return;
    const secondConfirm = confirm(`🚨 KONFIRMASI AKHIR\n\nTindakan ini TIDAK BISA dibatalkan. Seluruh ${total} data santri akan hilang selamanya.\n\nKetuk OK untuk menghapus semua data santri sekarang.`);
    if (!secondConfirm) return;
    setIsBulkProcessing(true);
    setBulkProgressTotal(total);
    setBulkProgressCurrent(0);
    setBulkProgressText('Memulai penghapusan...');
    await new Promise(r => setTimeout(r, 80));
    try {
      const students = [...data.students];
      for (let i = 0; i < students.length; i++) {
        const st = students[i];
        setBulkProgressText(`Menghapus: ${st.nama || `Santri ke-${i+1}`}`);
        setBulkProgressCurrent(i + 1);
        // eslint-disable-next-line no-await-in-loop
        await deleteFromDb('students', st.id);
      }
      setBulkProgressText('Selesai!');
      showNotification(`Berhasil menghapus semua ${total} data santri.`);
    } catch (err) {
      console.error(err);
      showNotification('Gagal menghapus semua data santri: ' + err.message, 'error');
    } finally {
      setIsBulkProcessing(false);
      setBulkProgressText('');
      setBulkProgressCurrent(0);
      setBulkProgressTotal(0);
    }
  };

  const renderFullTable = () => {
    switch (activeTab) {
      case 'backup_restore':
        return <BackupRestorePanel />;
      case 'settings': {
        const filteredSettings = hideInactive ? sortedData.filter(s => s.isActive) : sortedData;
        return (
          <div>
            <div className="flex items-center gap-2 mb-3 pb-3 border-b">
              <label className="flex items-center gap-2 text-sm cursor-pointer select-none bg-gray-50 hover:bg-gray-100 px-3 py-1.5 rounded-lg border">
                <input type="checkbox" checked={hideInactive} onChange={e => setHideInactive(e.target.checked)} className="w-4 h-4 text-emerald-600 rounded" />
                Sembunyikan semester nonaktif
              </label>
              <span className="text-xs text-gray-400">({filteredSettings.length} dari {sortedData.length} ditampilkan)</span>
            </div>
            <table className="w-full text-left border-collapse">
              <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                  <SortableHeader label="Tahun" sortKey="tahun" />
                  <SortableHeader label="Tahun Arab" sortKey="tahun_arab" />
                  <SortableHeader label="Semester" sortKey="semester" />
                  <SortableHeader label="Semester Arab" sortKey="semester_arab" />
                  <SortableHeader label="Status" sortKey="isActive" className="text-center" />
                  <th className="p-3 border-b text-center">Aksi</th>
              </tr></thead>
              <tbody>{filteredSettings.map(s => (
                  <tr key={s.id} className={`border-b hover:bg-gray-50 ${s.isActive ? 'bg-emerald-50' : ''}`}><td className="p-3 font-semibold">{s.tahun}</td><td className="p-3 font-arabic" dir="rtl">{s.tahun_arab}</td><td className="p-3">{s.semester}</td><td className="p-3 font-arabic" dir="rtl">{s.semester_arab}</td><td className="p-3 text-center">
                      <label className="relative inline-flex items-center cursor-pointer">
                        <input type="checkbox" className="sr-only peer" checked={s.isActive} onChange={(e) => saveToDb('settings', s.id, { ...s, isActive: e.target.checked }, true, `Mengubah status semester ${s.tahun} ${s.semester}`)} />
                        <div className="w-11 h-6 bg-gray-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-emerald-600"></div>
                      </label>
                  </td><td className="p-3 text-center"><button onClick={() => handleOpenModal(s)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('settings', s.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
                ))}</tbody>
            </table>
          </div>
        );
      }
      case 'teachers':
        return (
          <table className="w-full text-left border-collapse">
            <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                <SortableHeader label="Nama Guru" sortKey="nama" />
                <SortableHeader label="NIP / Identitas" sortKey="nip" />
                <SortableHeader label="Posisi / Jabatan" sortKey="posisi" />
                <th className="p-3 border-b text-center">Aksi</th>
            </tr></thead>
            <tbody>{sortedData.map(t => (
                <tr key={t.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{t.nama}</td><td className="p-3 text-gray-600">{t.nip || '-'}</td><td className="p-3 text-gray-600">{t.posisi || '-'}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(t)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('teachers', t.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
              ))}</tbody>
          </table>
        );
      case 'subjectCategories':
        return (
          <table className="w-full text-left border-collapse">
            <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                <SortableHeader label="Kategori Pelajaran (ID)" sortKey="name" />
                <SortableHeader label="Kategori Pelajaran (AR)" sortKey="nameAr" className="text-right" />
                <th className="p-3 border-b text-center">Aksi</th>
            </tr></thead>
            <tbody>{sortedData.map(c => (
                <tr key={c.id} className="border-b hover:bg-gray-50">
                  <td className="p-3 font-semibold">{c.name}</td>
                  <td className="p-3 text-right font-arabic" dir="rtl">{c.nameAr}</td>
                  <td className="p-3 text-center">
                    <button onClick={() => handleOpenModal(c)} className="text-blue-500 p-1"><Edit2 size={16}/></button>
                    <button onClick={() => deleteFromDb('subjectCategories', c.id)} className="text-red-500 p-1"><Trash2 size={16}/></button>
                  </td>
                </tr>
              ))}</tbody>
          </table>
        );
      case 'masterSubjects':
        return (
          <table className="w-full text-left border-collapse">
            <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                <SortableHeader label="Pelajaran Utama (Indo)" sortKey="nameId" />
                <SortableHeader label="Pelajaran Utama (Arab)" sortKey="nameAr" className="text-right" />
                <th className="p-3 border-b text-center">Aksi</th>
            </tr></thead>
            <tbody>{sortedData.map(m => (
                <tr key={m.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{m.nameId}</td><td className="p-3 text-right font-arabic" dir="rtl">{m.nameAr}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(m)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('masterSubjects', m.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
              ))}</tbody>
          </table>
        );
      case 'subjects':
        return (
          <table className="w-full text-left border-collapse">
            <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                <th className="p-3 border-b text-center">No.</th>
                <SortableHeader label="Kelas" sortKey="kelas" className="text-center" />
                <SortableHeader label="Kategori" sortKey="kategori" />
                <SortableHeader label="Mapel (ID)" sortKey="nameId" />
                <SortableHeader label="Mapel (AR)" sortKey="nameAr" className="text-right" />
                <SortableHeader label="KKM & Guru" sortKey="kkm" className="text-center" />
                <th className="p-3 border-b text-center">Aksi</th>
            </tr></thead>
            <tbody>{groupedSubjects.map((row, index) => {
                  if (row.type === 'group') {
                      return (
                          <tr key={`group-${row.kelas}-${index}`} className="bg-emerald-50">
                              <td colSpan="7" className="p-3 font-semibold text-emerald-800">Kelas: {row.kelas}</td>
                          </tr>
                      );
                  }
                  const sub = row.subject;
                  return (
                      <tr key={sub.id} className="border-b hover:bg-gray-50">
                          <td className="p-3 text-center font-semibold text-gray-700">{row.number}</td>
                          <td className="p-3 text-center font-semibold text-gray-800">{getSubjectClassLabel(sub, allData?.classes || data.classes) || 'Semua'}</td>
                          <td className="p-3"><div className="text-xs text-emerald-700 font-bold">{sub.kategori || '-'}</div></td>
                          <td className="p-3 font-semibold">{sub.nameId}</td>
                          <td className="p-3 text-right font-arabic" dir="rtl">{sub.nameAr}</td>
                          <td className="p-3 text-center"><div className="font-bold text-yellow-600">{sub.kkm}</div><div className="text-[11px] text-gray-500 mt-1">{sub.guru || '-'}</div></td>
                          <td className="p-3 text-center"><button onClick={() => handleOpenModal(sub)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('subjects', sub.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td>
                      </tr>
                  );
              })}</tbody>
          </table>
        );
            case 'presences':
                return (
                    <table className="w-full text-left border-collapse">
                        <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                                <SortableHeader label="Nama Aspek Presensi" sortKey="name" />
                                <SortableHeader label="Nama (AR)" sortKey="nameAr" className="text-right" />
                                <th className="p-3 border-b text-center">Aksi</th>
                        </tr></thead>
                        <tbody>{sortedData.map(p => (
                                <tr key={p.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{p.name}</td><td className="p-3 text-right font-arabic" dir="rtl">{p.nameAr}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(p)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('presences', p.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
                            ))}</tbody>
                    </table>
                );
      case 'characterTraits':
        return (
          <table className="w-full text-left border-collapse">
            <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                <SortableHeader label="Aspek Sikap/Kesantrian (ID)" sortKey="name" />
                <SortableHeader label="Nama (AR)" sortKey="nameAr" className="text-right" />
                <th className="p-3 border-b text-center">Aksi</th>
            </tr></thead>
            <tbody>{sortedData.map(p => (
                <tr key={p.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{p.name}</td><td className="p-3 text-right font-arabic" dir="rtl">{p.nameAr}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(p)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('characterTraits', p.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
              ))}</tbody>
          </table>
        );
      case 'extracurriculars':
        return (
          <table className="w-full text-left border-collapse">
            <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                <SortableHeader label="Nama Ekstrakurikuler (ID)" sortKey="name" />
                <SortableHeader label="Nama (AR)" sortKey="nameAr" className="text-right" />
                <th className="p-3 border-b text-center">Aksi</th>
            </tr></thead>
            <tbody>{sortedData.map(p => (
                <tr key={p.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{p.name}</td><td className="p-3 text-right font-arabic" dir="rtl">{p.nameAr}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(p)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('extracurriculars', p.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
              ))}</tbody>
          </table>
        );
      case 'fonts':
        return (
          <div className="space-y-4">
            <div className="bg-blue-50 text-blue-800 p-3 rounded-lg text-sm">
                <b>Penting:</b> Untuk font dari luar (misal Google Fonts), Anda WAJIB mengisi <b>URL Import</b> agar PDF bisa memuatnya. Contoh URL: <i>https://fonts.googleapis.com/css2?family=Oswald&display=swap</i>
            </div>
            <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                    <SortableHeader label="Nama Font" sortKey="name" />
                    <SortableHeader label="CSS Font-Family" sortKey="value" />
                    <th className="p-3 border-b text-center">Aksi</th>
                </tr></thead>
                <tbody>{sortedData.map(f => (
                    <tr key={f.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{f.name}</td><td className="p-3 text-gray-500 font-mono text-xs">{f.value}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(f)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('fonts', f.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
                ))}</tbody>
            </table>
          </div>
        );
            case 'students':
                return (
                        <div className="relative print:static">
                            <div className="mb-4 flex flex-wrap gap-2 items-center">
                                <button disabled={isBulkProcessing} onClick={() => generateExcelTemplate('students')} className="bg-emerald-600 text-white px-4 py-2 rounded-lg flex items-center gap-2 hover:bg-emerald-700 shadow-sm disabled:opacity-50"><Download size={16}/> Download Template Excel</button>
                                <label className={`bg-emerald-100 text-emerald-700 px-4 py-2 rounded-lg cursor-pointer flex items-center gap-2 hover:bg-emerald-200 ${isBulkProcessing ? 'opacity-50 pointer-events-none' : ''}`}>
                                    <Upload size={18} /> Impor Excel <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => handleImportExcel(e, 'students')} />
                                </label>
                                <button
                                    onClick={handleDeleteAllStudents}
                                    disabled={isBulkProcessing}
                                    className="ml-auto bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 shadow-sm transition font-semibold disabled:opacity-50"
                                >
                                    <Trash2 size={16}/> Hapus Semua Santri ({data.students?.length || 0})
                                </button>
                            </div>
                            <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                    <SortableHeader label="NIS" sortKey="nis" />
                    <SortableHeader label="Nama Santri" sortKey="nama" />
                    <SortableHeader label="Nama Arab" sortKey="nama_arab" />
                    <SortableHeader label="Kelas" sortKey="kelas" />
                    <th className="p-3 border-b text-center">Aksi</th>
                </tr></thead>
                <tbody>{sortedData.map(st => (<tr key={st.id} className="border-b hover:bg-gray-50"><td className="p-3">{st.nis}</td><td className="p-3 font-semibold">{st.nama}</td><td className="p-3 font-arabic" dir="rtl">{st.nama_arab}</td><td className="p-3">{getClassNameFromValue(allData?.classes || data.classes, st.kelas)}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(st)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('students', st.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>))}</tbody>
              </table>

              {/* OVERLAY PROGRESS SANTRI */}
              {isBulkProcessing && (
                <div className="fixed inset-0 bg-gray-900/80 backdrop-blur-sm z-[9999] flex flex-col items-center justify-center pointer-events-auto">
                  <div className="bg-white p-8 rounded-2xl shadow-2xl flex flex-col items-center w-full max-w-md mx-4 text-center">
                    <div className="animate-spin rounded-full h-14 w-14 border-4 border-emerald-100 border-b-emerald-600 mb-5"></div>
                    <h3 className="text-lg font-bold text-gray-800 mb-1">Sistem Terkunci Sementara</h3>
                    <p className="text-emerald-600 font-medium text-sm mb-4">{bulkProgressText}</p>
                    {bulkProgressTotal > 0 && (
                      <div className="w-full">
                        <div className="flex justify-between text-xs text-gray-500 mb-1">
                          <span>{bulkProgressCurrent} dari {bulkProgressTotal}</span>
                          <span>{Math.round((bulkProgressCurrent / bulkProgressTotal) * 100)}%</span>
                        </div>
                        <div className="w-full bg-gray-100 rounded-full h-3 overflow-hidden">
                          <div
                            className="bg-emerald-500 h-3 rounded-full transition-all duration-300"
                            style={{ width: `${Math.round((bulkProgressCurrent / bulkProgressTotal) * 100)}%` }}
                          />
                        </div>
                      </div>
                    )}
                    <p className="text-xs text-gray-400 mt-4">Mohon jangan tutup halaman ini selama proses berlangsung.</p>
                  </div>
                </div>
              )}
            </div>
        );
      case 'studentFields':
        return (
            <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                    <SortableHeader label="Nama Label" sortKey="name" />
                    <SortableHeader label="Variabel (Key)" sortKey="key" />
                    <th className="p-3 border-b text-center">Aksi</th>
                </tr></thead>
                <tbody>{sortedData.map(f => (<tr key={f.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{f.name}</td><td className="p-3 font-mono text-sm text-gray-500">{`{{${f.key}}}`}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(f)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('studentFields', f.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>))}</tbody>
            </table>
        );
      case 'classes':
        return (
            <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                    <SortableHeader label="Kelas" sortKey="name" />
                    <SortableHeader label="Kelas Arab" sortKey="name_arab" />
                    <SortableHeader label="Wali Kelas" sortKey="wali" />
                    <th className="p-3 border-b text-center">Aksi</th>
                </tr></thead>
                <tbody>{sortedData.map(c => (<tr key={c.id} className="border-b hover:bg-gray-50"><td className="p-3">{c.name}</td><td className="p-3 font-arabic" dir="rtl">{c.name_arab}</td><td className="p-3">{c.wali}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(c)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('classes', c.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>))}</tbody>
            </table>
        );
      case 'users':
        return (
            <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                    <SortableHeader label="Nama" sortKey="name" />
                    <SortableHeader label="Username" sortKey="username" />
                    <SortableHeader label="Role" sortKey="role" />
                    <th className="p-3 border-b text-center">Aksi</th>
                </tr></thead>
                <tbody>{sortedData.map(u => (<tr key={u.id} className="border-b hover:bg-gray-50"><td className="p-3">{u.name}</td><td className="p-3">{u.username}</td><td className="p-3">{u.role}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(u)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('users', u.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>))}</tbody>
            </table>
        );
      default: return null;
    }
  };

  const renderForm = () => {
    switch(activeTab) {
        case 'settings': return (
            <div className="space-y-4">
                <div>
                    <input 
                        className="w-full p-2 border rounded" 
                        placeholder="Tahun Ajaran (Misal: 2024/2025)" 
                        value={formData.tahun || ''} 
                        onChange={e => setFormData({...formData, tahun: e.target.value})}
                        onBlur={async (e) => {
                            if (!e.target.value) return;
                            const translated = await translateToArabic(e.target.value);
                            const arabicDigits = convertLatinDigitsToArabic(translated || e.target.value);
                            setFormData(prev => ({...prev, tahun_arab: arabicDigits}));
                        }}
                    />
                </div>
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Tahun Arab (السنة) - Terisi Otomatis" dir="rtl" value={formData.tahun_arab || ''} onChange={e => setFormData({...formData, tahun_arab: e.target.value})} />
                
                {/* Sembunyikan field semester jika sedang membuat Tahun Ajaran Baru */}
                {(editingItem && editingItem.tahun) && (
                    <>
                        <div>
                            <select 
                                className="w-full p-2 border rounded" 
                                value={formData.semester || 'Ganjil'} 
                                onChange={async (e) => {
                                    const semester = e.target.value;
                                    setFormData(prev => ({...prev, semester}));
                                    if (!semester) return;
                                    const translated = await translateToArabic(semester);
                                    setFormData(prev => ({...prev, semester_arab: translated || semester}));
                                }}
                            >
                                <option value="Ganjil">Ganjil</option>
                                <option value="Genap">Genap</option>
                            </select>
                        </div>
                        <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Semester Arab (الفصل) - Terisi Otomatis" dir="rtl" value={formData.semester_arab || ''} onChange={e => setFormData({...formData, semester_arab: e.target.value})} />
                    </>
                )}
                <label className="flex items-center gap-2"><input type="checkbox" checked={formData.isActive || false} onChange={e => setFormData({...formData, isActive: e.target.checked})} /> Jadikan Aktif</label>
                <p className="text-[10px] text-gray-500 italic">*Kolom Arab akan otomatis terisi menggunakan Google Translate saat Anda selesai mengetik atau memilih.</p>
            </div>
        );
        case 'teachers': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Nama Lengkap Guru" value={formData.nama || ''} onChange={e => setFormData({...formData, nama: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="NIP / NUPTK (Opsional)" value={formData.nip || ''} onChange={e => setFormData({...formData, nip: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="Posisi/Jabatan (Opsional)" value={formData.posisi || ''} onChange={e => setFormData({...formData, posisi: e.target.value})} />
            </div>
        );
        case 'subjectCategories': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded font-semibold" placeholder="Kategori Pelajaran (ID)" value={formData.name || ''} 
                    onChange={e => setFormData({...formData, name: e.target.value})} 
                    onBlur={async (e) => {
                        if (e.target.value && !formData.nameAr) {
                            const translated = await translateToArabic(e.target.value);
                            setFormData(prev => ({...prev, nameAr: translated}));
                        }
                    }}
                />
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Kategori (Arab) - Terisi Otomatis" dir="rtl" value={formData.nameAr || ''} onChange={e => setFormData({...formData, nameAr: e.target.value})} />
                <p className="text-[10px] text-gray-500 italic">*Ketik nama kategori (Indonesia) dan klik sembarang di luar kotak untuk Google Translate otomatis.</p>
            </div>
        );
        case 'masterSubjects': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded font-semibold" placeholder="Nama Pelajaran (Indo)" value={formData.nameId || ''} 
                    onChange={e => setFormData({...formData, nameId: e.target.value})} 
                    onBlur={async (e) => {
                        if (e.target.value && !formData.nameAr) {
                            const translated = await translateToArabic(e.target.value);
                            setFormData(prev => ({...prev, nameAr: translated}));
                        }
                    }}
                />
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Nama Pelajaran (Arab) - Terisi Otomatis" dir="rtl" value={formData.nameAr || ''} onChange={e => setFormData({...formData, nameAr: e.target.value})} />
                <p className="text-[10px] text-gray-500 italic">*Ketik nama pelajaran (Indonesia) dan klik sembarang di luar kotak. Kolom Arab akan otomatis diterjemahkan menggunakan Google Translate.</p>
            </div>
        );
        case 'subjects': return (
            <div className="space-y-4">
                <select className="w-full p-2 border rounded bg-yellow-50 font-bold" value={formData.kategori || ''} onChange={e => setFormData({...formData, kategori: e.target.value})}>
                    <option value="">-- Pilih Kategori Pelajaran --</option>
                    {(data.subjectCategories || []).map(c => <option key={c.id} value={c.name}>{c.name}</option>)}
                </select>
                <select className="w-full p-2 border rounded font-bold text-emerald-800" multiple value={Array.isArray(formData.kelas) ? formData.kelas : formData.kelas ? [formData.kelas] : []} onChange={e => {
                    const selected = Array.from(e.target.selectedOptions).map(o => o.value);
                    setFormData({...formData, kelas: selected});
                }}>
                    {data.classes.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                </select>
                <p className="text-[10px] text-gray-500 italic">*Pilih beberapa kelas jika pelajaran ini berlaku untuk lebih dari satu kelas. Biarkan kosong untuk semua kelas.</p>
                <select className="w-full p-2 border rounded font-bold text-blue-800" value={formData.nameId || ''} onChange={e => {
                    const selectedMaster = (data.masterSubjects || []).find(m => m.nameId === e.target.value);
                    setFormData({...formData, nameId: e.target.value, nameAr: selectedMaster ? selectedMaster.nameAr : formData.nameAr});
                }}>
                    <option value="">-- Pilih Pelajaran --</option>
                    {(data.masterSubjects || []).map(m => <option key={m.id} value={m.nameId}>{m.nameId}</option>)}
                </select>
                <input className="w-full p-2 border rounded text-right font-arabic bg-gray-100" placeholder="Nama Pelajaran (Arab)" dir="rtl" value={formData.nameAr || ''} disabled title="Otomatis mengikuti pilihan pelajaran di atas" />
                <select className="w-full p-2 border rounded font-bold text-purple-800" value={formData.guru || ''} onChange={e => setFormData({...formData, guru: e.target.value})}>
                    <option value="">-- Pilih Guru Pengampu --</option>
                    {data.teachers.map(t => <option key={t.id} value={t.nama}>{t.nama}</option>)}
                </select>
                <input type="number" className="w-full p-2 border rounded" placeholder="Standar KKM" value={formData.kkm || ''} onChange={e => setFormData({...formData, kkm: e.target.value})} />
                <label className="flex items-center gap-2"><input type="checkbox" checked={formData.isIjazah || false} onChange={e => setFormData({...formData, isIjazah: e.target.checked})} /> Tampilkan di Ijazah default</label>
            </div>
        );
        case 'presences': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Nama Aspek (Misal: Sakit, Izin, Tanpa Keterangan)" value={formData.name || ''} onChange={e => setFormData({...formData, name: e.target.value})}
                    onBlur={async (e) => {
                        if (e.target.value && !formData.nameAr) {
                            const translated = await translateToArabic(e.target.value);
                            setFormData(prev => ({...prev, nameAr: translated}));
                        }
                    }}
                />
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Nama Arab (Terisi Otomatis Google Translate)" dir="rtl" value={formData.nameAr || ''} onChange={e => setFormData({...formData, nameAr: e.target.value})} />
                <p className="text-[10px] text-gray-500 italic">*Ketik nama aspek (Indonesia) lalu klik di luar kotak untuk Google Translate otomatis.</p>
            </div>
        );
        case 'characterTraits': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Aspek Kesantrian/Sikap (Misal: Kedisiplinan)" value={formData.name || ''} 
                    onChange={e => setFormData({...formData, name: e.target.value})} 
                    onBlur={async (e) => {
                        if (e.target.value && !formData.nameAr) {
                            const translated = await translateToArabic(e.target.value);
                            setFormData(prev => ({...prev, nameAr: translated}));
                        }
                    }}
                />
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Nama Arab (Terisi Otomatis Google Translate)" dir="rtl" value={formData.nameAr || ''} onChange={e => setFormData({...formData, nameAr: e.target.value})} />
                <p className="text-[10px] text-gray-500 italic">*Ketik nama aspek (Indonesia) lalu klik di luar kotak untuk Google Translate otomatis.</p>
            </div>
        );
        case 'extracurriculars': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Nama Ekstrakurikuler (Misal: Komputer)" value={formData.name || ''} 
                    onChange={e => setFormData({...formData, name: e.target.value})} 
                    onBlur={async (e) => {
                        if (e.target.value && !formData.nameAr) {
                            const translated = await translateToArabic(e.target.value);
                            setFormData(prev => ({...prev, nameAr: translated}));
                        }
                    }}
                />
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Nama Arab (Terisi Otomatis Google Translate)" dir="rtl" value={formData.nameAr || ''} onChange={e => setFormData({...formData, nameAr: e.target.value})} />
                <p className="text-[10px] text-gray-500 italic">*Ketik nama ekskul (Indonesia) lalu klik di luar kotak untuk Google Translate otomatis.</p>
            </div>
        );
        case 'fonts': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Nama Font (Misal: Roboto)" value={formData.name || ''} onChange={e => setFormData({...formData, name: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="CSS Font-Family (Misal: 'Roboto', sans-serif)" value={formData.value || ''} onChange={e => setFormData({...formData, value: e.target.value})} />
                <input className="w-full p-2 border rounded text-sm text-blue-600 bg-blue-50" placeholder="URL Import Font (Misal: https://fonts.googleapis...)" value={formData.url || ''} onChange={e => setFormData({...formData, url: e.target.value})} />
            </div>
        );
        case 'studentFields': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Nama Kolom (Misal: Tempat Lahir)" value={formData.name || ''} onChange={e => setFormData({...formData, name: e.target.value, key: e.target.value.toLowerCase().replace(/[^a-z0-9]/g, '_')})} />
                <input className="w-full p-2 border rounded bg-gray-100" value={formData.key || ''} disabled placeholder="Otomatis menjadi key" />
            </div>
        );
        case 'students': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="NIS" value={formData.nis || ''} onChange={e => setFormData({...formData, nis: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="Nama Lengkap" value={formData.nama || ''} onChange={e => setFormData({...formData, nama: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="Nama Arab (النام)" value={formData.nama_arab || ''} onChange={e => setFormData({...formData, nama_arab: e.target.value})} />
                <select className="w-full p-2 border rounded" value={formData.kelas || ''} onChange={e => setFormData({...formData, kelas: e.target.value})}>
                    <option value="">Pilih Kelas</option>{data.classes.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                </select>
                {data.studentFields.map(f => (
                    <input key={f.key} className="w-full p-2 border rounded" placeholder={`Isi ${f.name}`} value={formData[f.key] || ''} onChange={e => setFormData({...formData, [f.key]: e.target.value})} />
                ))}
            </div>
        );
        case 'classes': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Nama Kelas" value={formData.name || ''} onChange={e => setFormData({...formData, name: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="Nama Kelas Arab (الفصل)" value={formData.name_arab || ''} onChange={e => setFormData({...formData, name_arab: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="Wali Kelas" value={formData.wali || ''} onChange={e => setFormData({...formData, wali: e.target.value})} />
            </div>
        );
        case 'users': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Nama" value={formData.name || ''} onChange={e => setFormData({...formData, name: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="Username" value={formData.username || ''} onChange={e => setFormData({...formData, username: e.target.value})} />
                <input className="w-full p-2 border rounded" type="password" placeholder="Password" value={formData.password || ''} onChange={e => setFormData({...formData, password: e.target.value})} />
                <select className="w-full p-2 border rounded" value={formData.role || 'user'} onChange={e => setFormData({...formData, role: e.target.value})}><option value="user">User</option><option value="admin">Admin</option></select>
            </div>
        );
        default: return null;
    }
  }

  const getTitle = () => {
    switch(activeTab) {
        case 'settings': return 'Tahun Ajaran & Semester';
        case 'teachers': return 'Guru Pengampu';
        case 'subjectCategories': return 'Kategori Pelajaran';
        case 'masterSubjects': return 'Daftar Pelajaran Utama';
        case 'subjects': return 'Plotting Pelajaran';
        case 'presences': return 'Aspek Presensi';
        case 'characterTraits': return 'Sikap & Kesantrian';
        case 'extracurriculars': return 'Ekstrakurikuler';
        case 'classes': return 'Daftar Kelas';
        case 'studentFields': return 'Field Form Santri';
        case 'students': return 'Data Santri';
        case 'fonts': return 'Pengaturan Font Kustom';
        case 'users': return 'Pengguna Sistem';
        case 'backup_restore': return 'Backup & Restore';
        default: return activeTab;
    }
  }

  return (
    <div className="bg-white rounded-xl shadow-sm p-6 border border-gray-100 flex flex-col h-[85vh]">
      <div className="mb-4 flex justify-between items-center shrink-0 border-b pb-4">
        <h3 className="text-xl font-bold text-gray-800 capitalize">Data {getTitle()}</h3>
        {activeTab !== 'backup_restore' && (
          <button onClick={() => handleOpenModal()} className="bg-emerald-600 text-white px-4 py-2 rounded-lg flex items-center gap-2 hover:bg-emerald-700 shadow-sm transition"><Plus size={18} /> Tambah Data</button>
        )}
      </div>
      <div className="flex-1 overflow-y-auto custom-scrollbar pr-2">{renderFullTable()}</div>
      
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={`Form ${getTitle()}`}>
        {renderForm()}
        <div className="mt-6 flex justify-between items-center bg-gray-50 p-3 rounded-lg border border-gray-200">
            <div className="flex items-center gap-4">
                <label className="flex items-center gap-2 text-sm">
                    <input type="checkbox" checked={autoSaveEnabled} onChange={e => setAutoSaveEnabled(e.target.checked)} />
                    <span>{autoSaveEnabled ? 'Autosave: Aktif (5s)' : 'Autosave: Nonaktif'}</span>
                </label>
                <div>
                    {isAutoSaving ? <span className="text-xs font-bold text-emerald-600 animate-pulse flex items-center gap-1"><Save size={14}/> Menyimpan otomatis...</span> : <span className="text-xs text-gray-500 font-medium">✅ Tersimpan aman di Cloud</span>}
                </div>
            </div>
            <div className="flex items-center gap-3">
                <button onClick={handleSave} className="bg-emerald-600 text-white px-4 py-2 rounded-lg font-medium hover:bg-emerald-700 transition shadow-sm">Simpan</button>
                <button onClick={() => setIsModalOpen(false)} className="bg-gray-800 text-white px-6 py-2 rounded-lg font-medium hover:bg-gray-900 transition shadow-sm">Tutup Form</button>
            </div>
        </div>
      </Modal>
    </div>
  );
};
// ==========================================
// RENDERER LOGIC FOR TABLES
// ==========================================
const groupBy = (array, key) => array.reduce((result, item) => {
    const group = item[key] || '';
    (result[group] = result[group] || []).push(item);
    return result;
}, {});

const renderDynamicTable = (el, data, studentGrades, classAverages = {}, isKatrol = false, mode = 'raport', classesData = []) => {
    const activeSetting = data.settings?.find(s => s.key === 'activeSetting')?.value || {};
    const toArabicNumbers = (val) => String(val).replace(/[0-9]/g, w => ['٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩'][w]);

    const columns = el.columns || [];
    if(columns.length === 0) return <div className="p-4 border bg-red-50 text-red-500 text-xs">Tabel belum dikonfigurasi. Silakan edit kolom di panel kiri.</div>;

    const toArabic = (val) => el.isRtl && val != null ? toArabicNumbers(val) : (val != null ? val : '-');

    const renderHeaders = () => (
        <thead>
            <tr>
                {columns.map((col, idx) => {
                    if (col.type === 'SPASI_KOSONG') {
                        return <th key={idx} style={{width: `${col.width}%`, height: col.height ? `${col.height}px` : 'auto', border: 'none', background: 'transparent'}}>{col.header === 'Kolom Baru' ? '' : col.header}</th>;
                    }
                    return (
                        <th key={idx} className="bg-gray-100 border border-black p-1 text-center font-bold" style={{width: `${col.width}%`, height: col.height ? `${col.height}px` : 'auto'}}>
                            {toArabic(col.header)}
                        </th>
                    );
                })}
            </tr>
        </thead>
    );

    const renderRowCells = (sub, idx) => {
        return columns.map((col, cIdx) => {
            let content = '-';
            let style = {};
            
            let gradeObj = studentGrades[sub.id];
            let rawGrade = 0;
            if (gradeObj && typeof gradeObj === 'object') {
                const r = computeRaportScore(gradeObj.uts, gradeObj.uas);
                if (r !== '') rawGrade = Number(r);
            } else if (gradeObj !== undefined && gradeObj !== '' && !isNaN(gradeObj)) {
                rawGrade = Number(gradeObj);
            }
            let finalGrade = isKatrol ? Math.max(rawGrade, Number(sub.kkm || 0)) : rawGrade;
            let isRed = !isKatrol && rawGrade > 0 && rawGrade < Number(sub.kkm || 0);

            switch(col.type) {
                case 'NO': content = toArabic(idx + 1); style={textAlign: 'center'}; break;
                case 'NO_AR': content = toArabicNumbers(idx + 1); style={textAlign: 'center', fontFamily: '"Amiri", "Scheherazade New", serif'}; break;
                
                case 'MAPEL_ID': content = toArabic(sub.nameId || sub.name); break;
                case 'MAPEL_AR': content = sub.nameAr || '-'; style={textAlign: 'right', fontFamily: '"Amiri", "Scheherazade New", serif'}; break;
                
                case 'KKM': content = toArabic(sub.kkm); style={textAlign: 'center'}; break;
                case 'KKM_AR': content = sub.kkm ? toArabicNumbers(sub.kkm) : '-'; style={textAlign: 'center', fontFamily: '"Amiri", "Scheherazade New", serif'}; break;
                
                case 'NILAI': 
                    content = finalGrade ? toArabic(finalGrade) : '-'; 
                    style={textAlign: 'center', fontWeight: 'bold', color: isRed ? 'red' : 'inherit'}; 
                    break;
                case 'NILAI_AR': 
                    content = finalGrade ? toArabicNumbers(finalGrade) : '-'; 
                    style={textAlign: 'center', fontWeight: 'bold', color: isRed ? 'red' : 'inherit', fontFamily: '"Amiri", "Scheherazade New", serif'}; 
                    break;
                    
                case 'RATA_KELAS': 
                    content = classAverages[sub.id] ? toArabic(classAverages[sub.id]) : '-'; 
                    style={textAlign: 'center'}; 
                    break;
                case 'RATA_KELAS_AR': 
                    content = classAverages[sub.id] ? toArabicNumbers(classAverages[sub.id]) : '-'; 
                    style={textAlign: 'center', fontFamily: '"Amiri", "Scheherazade New", serif'}; 
                    break;
                    
                case 'KATEGORI':
                    const catName = el.isRtl ? (data.subjectCategories?.find(c => normalizeValue(c.name) === normalizeValue(sub.kategori))?.nameAr || sub.kategori) : sub.kategori;
                    content = toArabic(catName);
                    style = el.isRtl ? { textAlign: 'right', fontFamily: '"Amiri", "Scheherazade New", serif' } : { textAlign: 'left' };
                    break;
                case 'KATEGORI_AR':
                    const catNameAr = data.subjectCategories?.find(c => normalizeValue(c.name) === normalizeValue(sub.kategori))?.nameAr || sub.kategori;
                    content = catNameAr;
                    style = { textAlign: 'right', fontFamily: '"Amiri", "Scheherazade New", serif' };
                    break;
                default: 
                    if(col.type.startsWith('PRESENCE_')) {
                        const pId = col.type.replace('PRESENCE_', '');
                        content = studentGrades[pId] ? toArabic(studentGrades[pId]) : '-';
                        style={textAlign: 'center'};
                    } else if(col.type.startsWith('SIKAP_')) {
                        const sId = col.type.replace('SIKAP_', '');
                        content = studentGrades[sId] ? toArabic(studentGrades[sId]) : '-';
                        style={textAlign: 'center'};
                    } else if(col.type.startsWith('EKSKUL_')) {
                        const eId = col.type.replace('EKSKUL_', '');
                        content = studentGrades[eId] ? toArabic(studentGrades[eId]) : '-';
                        style={textAlign: 'center'};
                    } else if(col.type === 'SPASI_KOSONG') {
                        content = '';
                    }
            }
            if (col.height) style.height = `${col.height}px`;
            
            if (col.type === 'SPASI_KOSONG') {
                return <td key={cIdx} style={{...style, border: 'none', background: 'transparent', padding: 0}}></td>;
            }
            return <td key={cIdx} className="border border-black p-1" style={style}>{content}</td>
        });
    };

    const baseSubjects = mode === 'ijazah' ? data.subjects.filter(s => s.isIjazah) : data.subjects;
    const subjectsToRender = el.filterClass
        ? filterSubjectsByClass(baseSubjects, el.filterClass, classesData.length ? classesData : data.classes || [])
        : baseSubjects;

    if (el.groupByCategory) {
        const grouped = groupBy(subjectsToRender, 'kategori');
        let globalIndex = 0;
        return (
            <table className="w-full border-collapse text-sm" dir={el.isRtl ? 'rtl' : 'ltr'} style={{ tableLayout: 'fixed', width: '100%', height: el.height ? `${el.height}px` : 'auto', fontSize: `${el.fontSize}px`, fontFamily: el.fontFamily }}>
                {renderHeaders()}
                <tbody>
                    {Object.entries(grouped).sort(([a],[b]) => a.localeCompare(b)).map(([cat, subs]) => {
                        // Jika RTL aktif, tampilkan nama kategori dalam bahasa Arab
                        const catLabel = el.isRtl
                            ? (data.subjectCategories?.find(c => normalizeValue(c.name) === normalizeValue(cat))?.nameAr || cat)
                            : cat;
                        return (
                        <React.Fragment key={cat}>
                            {cat && <tr><td colSpan={columns.length} className="border border-black p-1 font-bold bg-gray-50" style={{ textAlign: el.isRtl ? 'right' : 'left', fontFamily: el.isRtl ? '"Amiri", "Scheherazade New", serif' : 'inherit' }}>{catLabel}</td></tr>}
                            {subs.map(sub => {
                                globalIndex++;
                                return <tr key={sub.id}>{renderRowCells(sub, globalIndex - 1)}</tr>
                            })}
                        </React.Fragment>
                        );
                    })}
                </tbody>
            </table>
        );
    } else {
        return (
            <table className="w-full border-collapse text-sm" dir={el.isRtl ? 'rtl' : 'ltr'} style={{ tableLayout: 'fixed', width: '100%', height: el.height ? `${el.height}px` : 'auto', fontSize: `${el.fontSize}px`, fontFamily: el.fontFamily }}>
                {renderHeaders()}
                <tbody>
                    {subjectsToRender.map((sub, idx) => (
                        <tr key={sub.id}>{renderRowCells(sub, idx)}</tr>
                    ))}
                </tbody>
            </table>
        );
    }
};

// ==========================================
// DRAG & DROP LAYOUT BUILDER
// ==========================================
const defaultFontOptions = [
    { name: 'Arial (Bawaan)', value: 'Arial, sans-serif' },
    { name: 'Times New Roman (Klasik)', value: '"Times New Roman", Times, serif' },
    { name: 'Courier New (Mesin Tik)', value: '"Courier New", Courier, monospace' },
    { name: 'Amiri (Teks Arab)', value: '"Amiri", "Scheherazade New", serif' }
];

const pageDimensions = { 'A4': { width: 794, height: 1123 }, 'F4': { width: 816, height: 1248 } };
// A4 at 96 DPI: 794px = 21cm, so 1cm = 794/21 ≈ 37.8px
const PX_PER_CM = 794 / 21;
const defaultTableColumns = [
    { id: 'c1', header: 'No', type: 'NO', width: 5 },
    { id: 'c2', header: 'Mata Pelajaran', type: 'MAPEL_ID', width: 35 },
    { id: 'c3', header: 'المادة', type: 'MAPEL_AR', width: 35 },
    { id: 'c4', header: 'KKM', type: 'KKM', width: 10 },
    { id: 'c5', header: 'Nilai', type: 'NILAI', width: 15 }
];

const LayoutBuilder = () => {
    const { data, allData, activeSetting, saveToDb, deleteFromDb, showNotification, setAutoSaveStatus } = useContext(AppContext);
    const classesData = allData?.classes || data.classes || [];
    const [activeLayout, setActiveLayout] = useState(() => data.layouts?.length > 0 ? data.layouts[0].id : 'raport');
    const [elements, setElements] = useState([]);
    const [newLayoutName, setNewLayoutName] = useState('');
    const [showNewLayoutForm, setShowNewLayoutForm] = useState(false);
    const [pageSize, setPageSize] = useState('A4');
    const [guides, setGuides] = useState({ h: [], v: [] });
    const [selectedIds, setSelectedIds] = useState([]);
    const selectedElementId = selectedIds.length === 1 ? selectedIds[0] : null;
    const [selectionBox, setSelectionBox] = useState(null);
    const [currentPage, setCurrentPage] = useState(0);
    const [orientation, setOrientation] = useState('portrait');
    const [margins, setMargins] = useState({ top: 0, bottom: 0, left: 0, right: 0 });
    const [zoom, setZoom] = useState(0.5);
    const [past, setPast] = useState([]);
    const [future, setFuture] = useState([]);
    const [showToolbar, setShowToolbar] = useState(true);
    const [showRuler, setShowRuler] = useState(true);
    const [showGuideBars, setShowGuideBars] = useState(true);
    const [showGrid, setShowGrid] = useState(true);
    const [isBottomMenuExpanded, setIsBottomMenuExpanded] = useState(true);
    const canvasRef = useRef(null);
    const layoutContainerRef = useRef(null);
    const [isFullscreen, setIsFullscreen] = useState(false);
    const [showSidebar, setShowSidebar] = useState(true);

    useEffect(() => {
        const handleFullscreenChange = () => setIsFullscreen(!!document.fullscreenElement);
        document.addEventListener('fullscreenchange', handleFullscreenChange);
        return () => document.removeEventListener('fullscreenchange', handleFullscreenChange);
    }, []);

    const toggleFullscreen = async () => {
        try {
            if (!layoutContainerRef.current) return;
            if (!document.fullscreenElement) {
                await layoutContainerRef.current.requestFullscreen();
            } else {
                await document.exitFullscreen();
            }
        } catch (err) {
            console.error('Fullscreen error:', err);
            showNotification('Fullscreen tidak didukung di browser ini', 'error');
        }
    };

    const baseWidth = pageDimensions[pageSize].width;
    const baseHeight = pageDimensions[pageSize].height;
    const canvasWidth = orientation === 'landscape' ? baseHeight : baseWidth;
    const canvasHeight = orientation === 'landscape' ? baseWidth : baseHeight;

    const changePageSize = (newSize) => {
        const oldW = canvasWidth;
        const oldH = canvasHeight;
        const newDim = pageDimensions[newSize];
        const newW = orientation === 'landscape' ? newDim.height : newDim.width;
        const newH = orientation === 'landscape' ? newDim.width : newDim.height;
        const scaleX = newW / oldW;
        const scaleY = newH / oldH;
        setPast(p => [...p, elements]);
        setFuture([]);
        setElements(prev => prev.map(el => ({
            ...el,
            x: Math.round((el.x ?? 0) * scaleX),
            y: Math.round((el.y ?? 0) * scaleY),
            width: el.width ? Math.round(el.width * scaleX) : el.width,
            height: el.height ? Math.round(el.height * scaleY) : el.height,
        })));
        setPageSize(newSize);
    };

    const changeOrientation = (newOrientation) => {
        const oldW = canvasWidth;
        const oldH = canvasHeight;
        // Swap dimensions for new orientation
        const newW = oldH;
        const newH = oldW;
        const scaleX = newW / oldW;
        const scaleY = newH / oldH;
        setPast(p => [...p, elements]);
        setFuture([]);
        setElements(prev => prev.map(el => ({
            ...el,
            x: Math.round((el.x ?? 0) * scaleX),
            y: Math.round((el.y ?? 0) * scaleY),
            width: el.width ? Math.round(el.width * scaleX) : el.width,
            height: el.height ? Math.round(el.height * scaleY) : el.height,
        })));
        setOrientation(newOrientation);
    };
    
    const allFonts = useMemo(() => [...defaultFontOptions, ...(data.fonts || [])], [data.fonts]);

    useEffect(() => {
        const styleId = 'custom-fonts-style';
        let styleTag = document.getElementById(styleId);
        if (!styleTag) { styleTag = document.createElement('style'); styleTag.id = styleId; document.head.appendChild(styleTag); }
        const imports = data.fonts?.filter(f => f.url).map(f => `@import url('${f.url}');`).join('\n') || '';
        styleTag.innerHTML = imports;
    }, [data.fonts]);

    useEffect(() => {
        const handleKeyDown = (e) => {
            if (selectedIds.length === 0) return;
            
            // Jangan jalankan panah jika pengguna sedang mengetik di input/textarea
            if (document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA')) {
                return;
            }

            if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
                e.preventDefault();
                const step = e.shiftKey ? 10 : 1;
                let dx = 0, dy = 0;
                if (e.key === 'ArrowUp') dy = -step;
                if (e.key === 'ArrowDown') dy = step;
                if (e.key === 'ArrowLeft') dx = -step;
                if (e.key === 'ArrowRight') dx = step;
                // Geser semua elemen yang sedang dipilih (tanpa menambah history tiap pixel)
                setElements(prev => prev.map(el => {
                    if (!selectedIds.includes(el.id)) return el;
                    return { ...el, x: (el.x || 0) + dx, y: (el.y || 0) + dy };
                }));
            }
        };

        window.addEventListener('keydown', handleKeyDown);
        return () => window.removeEventListener('keydown', handleKeyDown);
    }, [selectedIds, elements]);

    const prevLayoutRef = useRef(null);

    useEffect(() => {
        if (prevLayoutRef.current === activeLayout) return;
        prevLayoutRef.current = activeLayout;

        const savedLayout = data.layouts.find(l => l.id === activeLayout);
        if (savedLayout) {
            setElements(savedLayout.elements || []);
            setPageSize(savedLayout.pageSize || 'A4');
            setOrientation(savedLayout.orientation || 'portrait');
            setGuides(savedLayout.guides || { h: [], v: [] });
            setMargins(savedLayout.margins || { top: 0, bottom: 0, left: 0, right: 0 });
        } else {
            setElements([]); setPageSize('A4'); setOrientation('portrait'); setGuides({ h: [], v: [] });
            setMargins({ top: 0, bottom: 0, left: 0, right: 0 });
        }
        setPast([]);
        setFuture([]);
        setSelectedIds([]);
        setCurrentPage(0);
    }, [activeLayout, data.layouts]);

    // Auto-save effect
    useEffect(() => {
        if (!activeLayout) return;
        
        // Cek apakah layout ini benar-benar ada di database
        // Jangan auto-save (membuat baru otomatis) jika layout sudah dihapus
        const layoutExists = data.layouts && data.layouts.some(l => l.id === activeLayout);
        if (!layoutExists && data.layouts?.length > 0) return;

        // Tampilkan status 'menyimpan' saat timer berjalan
        setAutoSaveStatus('saving');

        const timer = setTimeout(async () => {
            await saveToDb('layouts', activeLayout, {
                name: data.layouts?.find(l => l.id === activeLayout)?.name || activeLayout,
                elements,
                pageSize,
                orientation,
                guides,
                margins
            }, true); // silent save
            setAutoSaveStatus('saved');
            // Hilangkan indikator setelah 3 detik
            setTimeout(() => setAutoSaveStatus('idle'), 3000);
        }, 1500);
        return () => { clearTimeout(timer); };
    }, [elements, pageSize, orientation, guides, margins]);

    const addElement = (type, customKey = null) => {
        const isWatermark = type === 'watermark';
        const isLine = type === 'line';
        const isShape = type === 'shape';
        const elementType = isWatermark ? 'image' : type;

        let defaultContent = elementType === 'text' ? 'Teks Baru' : elementType === 'image' ? 'https://via.placeholder.com/150' : `{{${elementType}}}`;
        if (customKey) defaultContent = `{{${customKey}}}`;
        if (isLine) defaultContent = '';
        if (isShape) defaultContent = '';

        const newEl = {
            id: Date.now().toString(),
            pageIndex: currentPage,
            type: elementType, content: defaultContent,
            x: 50, y: 50, fontSize: 14, fontFamily: 'Arial, sans-serif', fontWeight: 'normal',
            width: isLine ? 400 : isShape ? 200 : (elementType === 'table_grades' ? 650 : elementType === 'image' ? (isWatermark ? 400 : 100) : 200),
            height: isLine ? 2 : isShape ? 50 : (elementType === 'table_grades' ? 300 : elementType === 'image' ? (isWatermark ? 400 : 100) : 30),
            zIndex: isWatermark ? 0 : 1,
            opacity: isWatermark ? 0.2 : 1,
            // line/shape specific
            ...(isLine ? { lineColor: '#000000', lineThickness: 2 } : {}),
            ...(isShape ? { shapeFill: '#000000', shapeRadius: 0, shapeBorder: 0, shapeBorderColor: '#000000' } : {}),
            ...(elementType === 'table_grades' ? { columns: [...defaultTableColumns], groupByCategory: false, filterClass: '' } : {})
        };
        setPast(p => [...p, elements]);
        setFuture([]);
        setElements([...elements, newEl]);
        setSelectedIds([newEl.id]);
    };

    const updateElement = (id, changes, commit = true) => {
        if (commit) {
            setPast(p => [...p, elements]);
            setFuture([]);
        }
        setElements(elements.map(el => el.id === id ? { ...el, ...changes } : el));
    };

    const removeElement = (id) => {
        setPast(p => [...p, elements]);
        setFuture([]);
        setElements(elements.filter(el => el.id !== id));
        setSelectedIds([]);
    };

    const duplicateElement = (id) => {
        const elToDuplicate = elements.find(el => el.id === id);
        if (!elToDuplicate) return;
        
        const newEl = {
            ...elToDuplicate,
            id: Date.now().toString(),
            x: elToDuplicate.x + 20,
            y: elToDuplicate.y + 20,
        };
        
        if (newEl.type === 'table_grades' && newEl.columns) {
            newEl.columns = JSON.parse(JSON.stringify(newEl.columns));
            newEl.columns.forEach((col, idx) => col.id = Date.now().toString() + '_' + idx);
        }
        
        setPast(p => [...p, elements]);
        setFuture([]);
        setElements([...elements, newEl]);
        setSelectedIds([newEl.id]);
    };

    const undo = () => {
        if (past.length === 0) return;
        const previous = past[past.length - 1];
        setPast(past.slice(0, past.length - 1));
        setFuture([elements, ...future]);
        setElements(previous);
        setSelectedIds([]);
    };

    const redo = () => {
        if (future.length === 0) return;
        const next = future[0];
        setFuture(future.slice(1));
        setPast([...past, elements]);
        setElements(next);
        setSelectedIds([]);
    };

    const saveLayout = () => saveToDb('layouts', activeLayout, { name: data.layouts.find(l => l.id === activeLayout)?.name || activeLayout, elements, pageSize, orientation, guides, margins }, false, `Menyimpan desain layout ${activeLayout}`);

    const createNewLayout = () => {
        if (!newLayoutName.trim()) {
            showNotification('Nama layout tidak boleh kosong', 'error');
            return;
        }
        const layoutId = newLayoutName.toLowerCase().replace(/\s+/g, '_');
        if (data.layouts.some(l => l.id === layoutId)) {
            showNotification('Layout dengan nama ini sudah ada', 'error');
            return;
        }
        saveToDb('layouts', layoutId, { name: newLayoutName, elements: [], pageSize: 'A4', orientation: 'portrait', guides: { v: [], h: [] }, margins: { top: 0, bottom: 0, left: 0, right: 0 } }, false, `Membuat layout baru: ${newLayoutName}`);
        setNewLayoutName('');
        setShowNewLayoutForm(false);
        setActiveLayout(layoutId);
    };

    const deleteLayout = (layoutId) => {
        if (confirm(`Hapus layout "${data.layouts.find(l => l.id === layoutId)?.name || layoutId}"?`)) {
            deleteFromDb('layouts', layoutId, false, `Menghapus layout ${layoutId}`);
            if (activeLayout === layoutId) {
                const remainingLayouts = data.layouts.filter(l => l.id !== layoutId);
                setActiveLayout(remainingLayouts.length > 0 ? remainingLayouts[0].id : 'raport');
            }
        }
    };

    const duplicateLayout = () => {
        const source = data.layouts.find(l => l.id === activeLayout);
        if (!source) return;
        const sourceName = source.name || activeLayout;
        const newName = `Salinan - ${sourceName}`;
        const newId = `salinan_${activeLayout}_${Date.now()}`;
        if (data.layouts.some(l => l.id === newId)) {
            showNotification('Gagal menduplikat layout, coba lagi.', 'error');
            return;
        }
        const clonedElements = JSON.parse(JSON.stringify(source.elements || []));
        // Assign new IDs to cloned elements to avoid conflicts
        clonedElements.forEach(el => { el.id = `${el.id}_copy_${Date.now()}`; });
        saveToDb('layouts', newId, {
            name: newName,
            elements: clonedElements,
            pageSize: source.pageSize || 'A4',
            orientation: source.orientation || 'portrait',
            guides: JSON.parse(JSON.stringify(source.guides || { h: [], v: [] })),
            margins: JSON.parse(JSON.stringify(source.margins || { top: 0, bottom: 0, left: 0, right: 0 }))
        }, false, `Menduplikat layout: ${sourceName}`);
        setActiveLayout(newId);
        showNotification(`Layout "${newName}" berhasil dibuat!`, 'success');
    };

    // Drag Logic
    const [draggingType, setDraggingType] = useState(null);
    const [dragIndex, setDragIndex] = useState(null);
    const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 });
    const [initialRect, setInitialRect] = useState(null);

    const handleResizeMouseDown = (e, el, direction) => {
        e.stopPropagation();
        setDraggingType(`resize_${direction}`);
        setDragIndex(el.id);
        const rect = canvasRef.current.getBoundingClientRect();
        const scaleX = canvasWidth / rect.width;
        const scaleY = canvasHeight / rect.height;
        let rawX = (e.clientX - rect.left) * scaleX; 
        let rawY = (e.clientY - rect.top) * scaleY;
        
        const domNode = document.querySelector(`[data-element-id="${el.id}"]`);
        const actualWidth = domNode ? domNode.offsetWidth : (el.width || (el.type === 'table_grades' ? 650 : 200));
        const actualHeight = domNode ? domNode.offsetHeight : (el.height || (el.type === 'table_grades' ? 300 : 30));

        setInitialRect({ 
            width: actualWidth, 
            height: actualHeight, 
            startX: rawX, 
            startY: rawY,
            elX: el.x,
            elY: el.y
        });
        setPast(p => [...p, elements]);
        setFuture([]);
    };

    const handleElementMouseDown = (e, el) => {
        e.stopPropagation();

        if (e.shiftKey) {
            // Shift+Klik: toggle elemen masuk/keluar dari seleksi
            setSelectedIds(prev => prev.includes(el.id) ? prev.filter(id => id !== el.id) : [...prev, el.id]);
            return;
        }

        // Jika klik pada elemen yang sudah ada di multi-seleksi → mulai multi-drag
        if (selectedIds.includes(el.id) && selectedIds.length > 1 && !el.locked) {
            const canvasRect = canvasRef.current.getBoundingClientRect();
            const scaleX = canvasWidth / canvasRect.width;
            const scaleY = canvasHeight / canvasRect.height;
            const mouseX = (e.clientX - canvasRect.left) * scaleX;
            const mouseY = (e.clientY - canvasRect.top) * scaleY;
            const positions = {};
            selectedIds.forEach(id => {
                const found = elements.find(el2 => el2.id === id);
                if (found) positions[id] = { x: found.x || 0, y: found.y || 0 };
            });
            setInitialRect({ mouseX, mouseY, positions });
            setDraggingType('multi_element');
            setDragIndex(el.id);
            setPast(p => [...p, elements]);
            setFuture([]);
            return;
        }

        // Single select + drag
        setSelectedIds([el.id]);
        if (!el.locked) {
            setDraggingType('element'); setDragIndex(el.id);
            const rect = e.target.getBoundingClientRect();
            setDragOffset({ x: e.clientX - rect.left, y: e.clientY - rect.top });
            setPast(p => [...p, elements]);
            setFuture([]);
        }
    };

    const startDragGuide = (e, type, index) => { e.stopPropagation(); setDraggingType(`guide_${type}`); setDragIndex(index); };
    const createHGuide = (e) => {
        e.stopPropagation(); const rect = canvasRef.current.getBoundingClientRect();
        const y = (e.clientY - rect.top) * (canvasHeight / rect.height);
        setGuides({ ...guides, h: [...guides.h, y] }); setDraggingType('guide_h'); setDragIndex(guides.h.length);
    };
    const createVGuide = (e) => {
        e.stopPropagation(); const rect = canvasRef.current.getBoundingClientRect();
        const x = (e.clientX - rect.left) * (canvasWidth / rect.width);
        setGuides({ ...guides, v: [...guides.v, x] }); setDraggingType('guide_v'); setDragIndex(guides.v.length);
    };

    const handleMouseMove = (e) => {
        if (!draggingType || !canvasRef.current) return;
        const canvasRect = canvasRef.current.getBoundingClientRect();
        const scaleX = canvasWidth / canvasRect.width; const scaleY = canvasHeight / canvasRect.height;
        let rawX = (e.clientX - canvasRect.left) * scaleX; let rawY = (e.clientY - canvasRect.top) * scaleY;

        if (draggingType === 'element') {
            let newX = rawX - dragOffset.x * scaleX; let newY = rawY - dragOffset.y * scaleY;
            const snapThreshold = 10; 
            guides.v.forEach(gx => { if (Math.abs(newX - gx) < snapThreshold) newX = gx; });
            guides.h.forEach(gy => { if (Math.abs(newY - gy) < snapThreshold) newY = gy; });
            updateElement(dragIndex, { x: newX, y: newY }, false);
        } else if (draggingType === 'multi_element') {
            if (!initialRect?.positions) return;
            const dx = rawX - initialRect.mouseX;
            const dy = rawY - initialRect.mouseY;
            setElements(prev => prev.map(el => {
                const initPos = initialRect.positions[el.id];
                if (!initPos) return el;
                return { ...el, x: initPos.x + dx, y: initPos.y + dy };
            }));
        } else if (draggingType === 'selection') {
            setSelectionBox(prev => prev ? { ...prev, endX: rawX, endY: rawY } : prev);
        } else if (draggingType && draggingType.startsWith('resize_')) {
            const activeEl = elements.find(el => el.id === dragIndex);
            if (!activeEl) return;
            const direction = draggingType.split('_')[1];
            
            let newWidth = initialRect.width;
            let newHeight = initialRect.height;
            let newX = initialRect.elX;
            let newY = initialRect.elY;
            
            const dx = rawX - initialRect.startX;
            const dy = rawY - initialRect.startY;
            
            if (direction.includes('e')) { newWidth = Math.max(10, initialRect.width + dx); }
            if (direction.includes('s')) { newHeight = Math.max(10, initialRect.height + dy); }
            if (direction.includes('w')) { 
                const deltaX = Math.min(initialRect.width - 10, dx);
                newX = initialRect.elX + deltaX; 
                newWidth = initialRect.width - deltaX; 
            }
            if (direction.includes('n')) { 
                const deltaY = Math.min(initialRect.height - 10, dy);
                newY = initialRect.elY + deltaY; 
                newHeight = initialRect.height - deltaY; 
            }
            
            updateElement(dragIndex, { width: newWidth, height: newHeight, x: newX, y: newY }, false);
        } else if (draggingType === 'guide_v') {
            const newGuides = { ...guides }; newGuides.v[dragIndex] = rawX; setGuides(newGuides);
        } else if (draggingType === 'guide_h') {
            const newGuides = { ...guides }; newGuides.h[dragIndex] = rawY; setGuides(newGuides);
        }
    };

    const handleMouseUp = () => {
        // Finalisasi rubber-band selection
        if (draggingType === 'selection' && selectionBox) {
            const minX = Math.min(selectionBox.startX, selectionBox.endX);
            const maxX = Math.max(selectionBox.startX, selectionBox.endX);
            const minY = Math.min(selectionBox.startY, selectionBox.endY);
            const maxY = Math.max(selectionBox.startY, selectionBox.endY);
            if (maxX - minX > 5 || maxY - minY > 5) {
                const selected = elements.filter(el => {
                    if ((el.pageIndex || 0) !== currentPage || el.locked) return false;
                    const ex = el.x || 0; const ey = el.y || 0;
                    const ew = el.width || 200; const eh = el.height || 30;
                    return ex < maxX && ex + ew > minX && ey < maxY && ey + eh > minY;
                });
                if (selected.length > 0) setSelectedIds(selected.map(el => el.id));
            }
            setSelectionBox(null);
        }
        if (draggingType?.startsWith('guide_')) {
            const newGuides = { ...guides };
            if (draggingType === 'guide_v' && (guides.v[dragIndex] < -20 || guides.v[dragIndex] > canvasWidth + 20)) newGuides.v.splice(dragIndex, 1);
            if (draggingType === 'guide_h' && (guides.h[dragIndex] < -20 || guides.h[dragIndex] > canvasHeight + 20)) newGuides.h.splice(dragIndex, 1);
            setGuides(newGuides);
        }
        setDraggingType(null); setDragIndex(null); setInitialRect(null);
    };

    // Klik pada background kanvas atau area scroll sekitarnya → mulai rubber-band selection
    const handleCanvasMouseDown = (e) => {
        // Hanya jalankan jika target bukan elemen interaktif atau dalam elemen kanvas
        if (
            e.target.closest('button') || 
            e.target.closest('input') || 
            e.target.closest('select') ||
            e.target.closest('[data-element-id]') ||
            e.target.closest('svg') // abaikan klik pada svg ruler
        ) return;

        const rect = canvasRef.current.getBoundingClientRect();
        const scaleX = canvasWidth / rect.width;
        const scaleY = canvasHeight / rect.height;
        const startX = (e.clientX - rect.left) * scaleX;
        const startY = (e.clientY - rect.top) * scaleY;
        if (!e.shiftKey) setSelectedIds([]);
        setSelectionBox({ startX, startY, endX: startX, endY: startY });
        setDraggingType('selection');
    };

    const handleImageUpload = async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        // Jika ukuran file < 200KB, boleh pakai base64 (aman untuk DB)
        // Jika lebih besar, upload ke Supabase Storage agar tidak gagal simpan
        const MAX_BASE64_SIZE = 200 * 1024; // 200KB

        if (file.size <= MAX_BASE64_SIZE) {
            const reader = new FileReader();
            reader.onload = (ev) => updateElement(selectedElementId, { content: ev.target.result });
            reader.readAsDataURL(file);
            return;
        }

        // Upload ke Supabase Storage
        showNotification('Mengupload gambar, mohon tunggu...', 'info');
        try {
            const ext = file.name.split('.').pop();
            const fileName = `layout_${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`;
            const { error: uploadError } = await supabase.storage
                .from('layout-images')
                .upload(fileName, file, { upsert: true, contentType: file.type });

            if (uploadError) throw uploadError;

            const { data: urlData } = supabase.storage
                .from('layout-images')
                .getPublicUrl(fileName);

            if (!urlData?.publicUrl) throw new Error('Gagal mendapatkan URL gambar');

            updateElement(selectedElementId, { content: urlData.publicUrl });
            showNotification('Gambar berhasil diupload!');
        } catch (err) {
            // Fallback: jika storage gagal (bucket belum dibuat), paksa base64
            showNotification('Storage tidak tersedia, menggunakan metode lain...', 'warning');
            const reader = new FileReader();
            reader.onload = (ev) => updateElement(selectedElementId, { content: ev.target.result });
            reader.readAsDataURL(file);
        }
    };

    const activeEl = elements.find(e => e.id === selectedElementId);

    const alignElement = (direction) => {
        if (!activeEl) return;
        
        // Coba dapatkan ukuran asli dari elemen di layar
        const domNode = document.querySelector(`[data-element-id="${selectedElementId}"]`);
        const actualWidth = domNode ? domNode.offsetWidth : (activeEl.width || 200);
        const actualHeight = domNode ? domNode.offsetHeight : (activeEl.height || 30);

        const changes = {
            'left':   { x: 0 },
            'center': { x: (canvasWidth - actualWidth) / 2 },
            'right':  { x: canvasWidth - actualWidth },
            'top':    { y: 0 },
            'middle': { y: (canvasHeight - actualHeight) / 2 },
            'bottom': { y: canvasHeight - actualHeight },
        }[direction];
        if (changes) updateElement(selectedElementId, changes);
    };

    // ---- MULTI-SELECT: Group, Ungroup, Align ----
    const groupElements = () => {
        if (selectedIds.length < 2) return;
        const selected = elements.filter(el => selectedIds.includes(el.id) && (el.pageIndex || 0) === currentPage);
        if (selected.length < 2) return;
        const minX = Math.min(...selected.map(el => el.x || 0));
        const minY = Math.min(...selected.map(el => el.y || 0));
        const maxX = Math.max(...selected.map(el => (el.x || 0) + (el.width || 200)));
        const maxY = Math.max(...selected.map(el => (el.y || 0) + (el.height || 30)));
        const children = selected.map(el => ({ ...el, x: (el.x || 0) - minX, y: (el.y || 0) - minY }));
        const groupEl = {
            id: Date.now().toString(), type: 'group', pageIndex: currentPage,
            x: minX, y: minY, width: maxX - minX, height: maxY - minY,
            zIndex: Math.max(...selected.map(el => el.zIndex ?? 1)), opacity: 1, children,
        };
        setPast(p => [...p, elements]); setFuture([]);
        setElements([...elements.filter(el => !selectedIds.includes(el.id)), groupEl]);
        setSelectedIds([groupEl.id]);
    };

    const ungroupElements = () => {
        const groupEl = elements.find(el => el.id === selectedElementId && el.type === 'group');
        if (!groupEl) return;
        const base = Date.now();
        const ungrouped = (groupEl.children || []).map((child, idx) => ({
            ...child,
            id: `${base}_${idx}_${Math.random().toString(36).slice(2)}`,
            x: (child.x || 0) + (groupEl.x || 0),
            y: (child.y || 0) + (groupEl.y || 0),
            pageIndex: currentPage,
        }));
        setPast(p => [...p, elements]); setFuture([]);
        setElements([...elements.filter(el => el.id !== selectedElementId), ...ungrouped]);
        setSelectedIds(ungrouped.map(el => el.id));
    };

    const alignMultiple = (direction) => {
        if (selectedIds.length < 2) return;
        const selected = elements.filter(el => selectedIds.includes(el.id));
        const updMap = {};
        if (direction === 'left') {
            const minX = Math.min(...selected.map(el => el.x || 0));
            selected.forEach(el => { updMap[el.id] = { x: minX }; });
        } else if (direction === 'right') {
            const maxX = Math.max(...selected.map(el => (el.x || 0) + (el.width || 200)));
            selected.forEach(el => { updMap[el.id] = { x: maxX - (el.width || 200) }; });
        } else if (direction === 'center') {
            const minX = Math.min(...selected.map(el => el.x || 0));
            const maxX = Math.max(...selected.map(el => (el.x || 0) + (el.width || 200)));
            const cx = (minX + maxX) / 2;
            selected.forEach(el => { updMap[el.id] = { x: cx - (el.width || 200) / 2 }; });
        } else if (direction === 'top') {
            const minY = Math.min(...selected.map(el => el.y || 0));
            selected.forEach(el => { updMap[el.id] = { y: minY }; });
        } else if (direction === 'bottom') {
            const maxY = Math.max(...selected.map(el => (el.y || 0) + (el.height || 30)));
            selected.forEach(el => { updMap[el.id] = { y: maxY - (el.height || 30) }; });
        } else if (direction === 'middle') {
            const minY = Math.min(...selected.map(el => el.y || 0));
            const maxY = Math.max(...selected.map(el => (el.y || 0) + (el.height || 30)));
            const cy = (minY + maxY) / 2;
            selected.forEach(el => { updMap[el.id] = { y: cy - (el.height || 30) / 2 }; });
        }
        setPast(p => [...p, elements]); setFuture([]);
        setElements(elements.map(el => selectedIds.includes(el.id) ? { ...el, ...(updMap[el.id] || {}) } : el));
    };
    
    // Memberikan objek dummy default agar renderDynamicTable tidak crash saat proses desain layout
    const mockStudentGrades = {};
    const mockClassAverages = {};

    return (
        <div ref={layoutContainerRef} className={`flex flex-col md:flex-row gap-6 print:h-auto print:block ${isFullscreen ? 'h-screen w-screen bg-gray-50 p-4' : 'h-[80vh]'}`}>
            {showSidebar && (
            <div className="w-full md:w-[320px] bg-white rounded-xl shadow-sm flex flex-col border border-gray-100 shrink-0 overflow-hidden print:hidden">
                <div className="p-4 border-b bg-gray-50 flex items-center justify-between z-10 shrink-0">
                    <h3 className="font-bold text-gray-800 text-lg">Layout Builder</h3>
                </div>
                
                <div className="flex-1 overflow-y-auto p-4 space-y-4 custom-scrollbar">
                    <div className="flex gap-2">
                        <select className="flex-1 p-2 border rounded-lg bg-white text-sm font-bold text-emerald-800" value={activeLayout} onChange={e => setActiveLayout(e.target.value)}>
                            {data.layouts && data.layouts.map(l => <option key={l.id} value={l.id}>{l.name || l.id}</option>)}
                        </select>
                        <button onClick={() => setShowNewLayoutForm(!showNewLayoutForm)} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 rounded-lg text-sm font-bold transition" title="Tambah layout baru"><Plus size={16}/></button>
                        <button onClick={duplicateLayout} className="bg-blue-500 hover:bg-blue-600 text-white px-3 rounded-lg text-sm font-bold transition" title="Duplikat layout ini"><Copy size={16}/></button>
                        {data.layouts && data.layouts.length > 1 && <button onClick={() => deleteLayout(activeLayout)} className="bg-red-500 hover:bg-red-600 text-white px-3 rounded-lg text-sm font-bold transition" title="Hapus layout"><Trash2 size={16}/></button>}
                    </div>
                    <div className="flex gap-2">
                        <select className="w-1/2 p-2 border rounded-lg bg-white text-sm font-bold text-blue-800" value={pageSize} onChange={e => changePageSize(e.target.value)}>
                            <option value="A4">Size: A4</option><option value="F4">Size: F4</option>
                        </select>
                        <select className="w-1/2 p-2 border rounded-lg bg-white text-sm font-bold text-blue-800" value={orientation} onChange={e => changeOrientation(e.target.value)}>
                            <option value="portrait">Potrait</option><option value="landscape">Landscape</option>
                        </select>
                    </div>
                    <div className="flex flex-col gap-2 p-3 bg-indigo-50 border border-indigo-100 rounded-lg">
                        <p className="text-xs font-bold text-indigo-800 uppercase tracking-wider">Batas Area Aman / Margin (mm)</p>
                        <p className="text-[10px] text-indigo-600 leading-tight">Garis merah putus-putus akan muncul di kanvas sebagai panduan area aman cetak (tidak akan ikut ter-print).</p>
                        <div className="grid grid-cols-4 gap-2 mt-1">
                            <div className="flex flex-col"><label className="text-[10px] text-indigo-600 mb-1 font-bold">Atas</label><input type="number" value={margins.top} onChange={e=>setMargins({...margins, top: Number(e.target.value)})} className="w-full p-1.5 border border-indigo-200 rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-indigo-600 mb-1 font-bold">Bawah</label><input type="number" value={margins.bottom} onChange={e=>setMargins({...margins, bottom: Number(e.target.value)})} className="w-full p-1.5 border border-indigo-200 rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-indigo-600 mb-1 font-bold">Kiri</label><input type="number" value={margins.left} onChange={e=>setMargins({...margins, left: Number(e.target.value)})} className="w-full p-1.5 border border-indigo-200 rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-indigo-600 mb-1 font-bold">Kanan</label><input type="number" value={margins.right} onChange={e=>setMargins({...margins, right: Number(e.target.value)})} className="w-full p-1.5 border border-indigo-200 rounded text-xs" /></div>
                        </div>
                    </div>
                    <button onClick={() => { setSelectedElementId(null); setTimeout(() => { document.title = "Print_Preview_Layout"; window.print(); }, 100); }} className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition shadow-sm"><Printer size={16}/> Print Preview</button>
                    {showNewLayoutForm && (
                        <div className="border-t pt-3 space-y-2">
                            <p className="text-xs font-semibold text-gray-600">Nama Layout Baru</p>
                            <input type="text" placeholder="Misal: Sertifikat, Izin, dll" className="w-full p-2 border rounded-lg text-sm" value={newLayoutName} onChange={e => setNewLayoutName(e.target.value)} onKeyPress={e => e.key === 'Enter' && createNewLayout()} />
                            <div className="flex gap-2">
                                <button onClick={createNewLayout} className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white py-2 rounded-lg text-sm font-bold transition">Buat</button>
                                <button onClick={() => {setShowNewLayoutForm(false); setNewLayoutName('');}} className="flex-1 bg-gray-300 hover:bg-gray-400 text-gray-700 py-2 rounded-lg text-sm font-bold transition">Batal</button>
                            </div>
                        </div>
                    )}
                    
                    <div className="space-y-2 border-b pb-4">
                        <p className="text-xs font-semibold text-gray-500 uppercase sticky top-0 bg-white z-10 pb-1">Tambah Elemen</p>
                        <button onClick={() => addElement('text')} className="w-full bg-gray-100 hover:bg-gray-200 text-gray-700 py-2 rounded text-sm flex items-center justify-center gap-2"><TypeIcon size={16}/> Teks Bebas</button>
                        <button onClick={() => addElement('image')} className="w-full bg-purple-50 hover:bg-purple-100 text-purple-700 py-2 rounded text-sm flex items-center justify-center gap-2"><ImageIcon size={16}/> Gambar (Logo/Stempel)</button>
                        <button onClick={() => addElement('watermark')} className="w-full bg-blue-50 hover:bg-blue-100 text-blue-700 py-2 rounded text-sm flex items-center justify-center gap-2"><ImageIcon size={16}/> Gambar Watermark</button>
                        <button onClick={() => addElement('table_grades')} className="w-full bg-orange-50 hover:bg-orange-100 text-orange-700 py-2 rounded text-sm flex items-center justify-center gap-2"><Columns size={16}/> Tabel Nilai Dinamis</button>
                        
                        <p className="text-xs font-semibold text-gray-500 uppercase mt-4 mb-1">Shape &amp; Garis</p>
                        <div className="grid grid-cols-2 gap-1">
                            <button onClick={() => addElement('line')} className="bg-gray-100 hover:bg-gray-200 text-gray-700 py-1.5 rounded text-sm flex items-center justify-center gap-2"><Minus size={14}/> Garis</button>
                            <button onClick={() => addElement('shape')} className="bg-gray-100 hover:bg-gray-200 text-gray-700 py-1.5 rounded text-sm flex items-center justify-center gap-2"><Square size={14}/> Kotak/Shape</button>
                        </div>
                        
                        <p className="text-xs font-semibold text-gray-500 uppercase mt-4 mb-1">Variabel Santri & Wali</p>
                        <div className="grid grid-cols-2 gap-1">
                            <button onClick={() => addElement('nama_santri')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><User size={14}/> Nama</button>
                            <button onClick={() => addElement('nama_santri_ar')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><User size={14}/> Nama (Arab)</button>
                            <button onClick={() => addElement('nis')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><CreditCard size={14}/> NIS</button>
                            <button onClick={() => addElement('kelas')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><BookOpen size={14}/> Kelas</button>
                            <button onClick={() => addElement('kelas_ar')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><BookOpen size={14}/> Kelas (Arab)</button>
                            <button onClick={() => addElement('wali_kelas')} className="bg-purple-50 hover:bg-purple-100 text-purple-700 py-1.5 rounded text-xs flex justify-center gap-1"><User size={14}/> Wali Kelas</button>
                            <button onClick={() => addElement('catatan_wali')} className="col-span-2 bg-pink-50 hover:bg-pink-100 text-pink-700 py-1.5 rounded text-xs flex justify-center gap-1"><FileSignature size={14}/> Catatan Wali Kelas</button>
                        </div>
                        
                        <p className="text-xs font-semibold text-gray-500 uppercase mt-4 mb-1">Variabel Umum</p>
                        <div className="grid grid-cols-2 gap-1">
                            <button onClick={() => addElement('tahun_ajaran')} className="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 py-1.5 rounded text-xs flex justify-center gap-1"><Calendar size={14}/> Tahun Ajaran</button>
                            <button onClick={() => addElement('tahun_ajaran_ar')} className="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 py-1.5 rounded text-xs flex justify-center gap-1"><Calendar size={14}/> Thn Ajaran (Arab)</button>
                            <button onClick={() => addElement('semester')} className="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 py-1.5 rounded text-xs flex justify-center gap-1"><BookOpen size={14}/> Semester</button>
                            <button onClick={() => addElement('semester_ar')} className="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 py-1.5 rounded text-xs flex justify-center gap-1"><BookOpen size={14}/> Semester (Arab)</button>
                        </div>
                        
                        {data.studentFields && data.studentFields.length > 0 && (
                        <div className="mt-2 grid grid-cols-2 gap-1">
                            {data.studentFields.map(f => (
                            <button key={f.id} onClick={() => addElement('custom', f.key)} className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 py-1.5 rounded text-xs flex justify-center gap-1 truncate px-1" title={f.name}>
                                <FileText size={14}/> {f.name}
                            </button>
                            ))}
                        </div>
                        )}

                        <div className="mt-6 pt-4 border-t">
                            <p className="text-[10px] font-bold text-gray-400 uppercase mb-2 tracking-wide">Daftar Lapisan (Layers)</p>
                            <div className="space-y-1 max-h-[150px] overflow-y-auto custom-scrollbar">
                                {elements.filter(el => (el.pageIndex || 0) === currentPage).map((el, i) => (
                                    <button 
                                        key={el.id} 
                                        onClick={(e) => {
                                            if (e.shiftKey) {
                                                setSelectedIds(prev => prev.includes(el.id) ? prev.filter(id => id !== el.id) : [...prev, el.id]);
                                            } else {
                                                setSelectedIds([el.id]);
                                            }
                                        }} 
                                        className={`w-full text-left px-2 py-1.5 rounded text-xs flex items-center justify-between transition ${selectedIds.includes(el.id) ? 'bg-blue-100 text-blue-800 font-bold border border-blue-200' : 'hover:bg-gray-100 text-gray-700 border border-transparent'}`}
                                    >
                                        <span className="truncate w-[80%]">{el.type === 'group' ? 'Grup Elemen' : el.type === 'image' ? (el.zIndex === 0 ? 'Gambar Watermark' : 'Gambar') : el.type === 'table_grades' ? 'Tabel Nilai' : (el.content || '').slice(0, 20) + ((el.content || '').length > 20 ? '...' : '')}</span>
                                        {el.locked && <Lock size={12} className="text-yellow-600"/>}
                                    </button>
                                ))}
                                {elements.filter(el => (el.pageIndex || 0) === currentPage).length === 0 && <p className="text-xs text-gray-400 italic">Belum ada elemen</p>}
                            </div>
                        </div>
                    </div>

                    {selectedIds.length > 1 && (
                        <div className="space-y-3 pt-2 pb-8 bg-indigo-50/50 p-3 rounded-lg border border-indigo-100">
                            <p className="text-xs font-bold text-indigo-800 uppercase flex items-center justify-between">
                                {selectedIds.length} Elemen Terpilih
                                <button onClick={() => setSelectedIds([])} className="text-gray-400 hover:text-gray-700"><X size={14}/></button>
                            </p>
                            
                            <div className="grid grid-cols-2 gap-2 mt-2">
                                <button onClick={groupElements} className="bg-white border hover:bg-gray-50 text-indigo-700 py-1.5 rounded text-[10px] font-bold transition flex items-center justify-center gap-1"><Layers size={14}/> Group</button>
                                <button onClick={() => {
                                    setPast(p => [...p, elements]); setFuture([]);
                                    setElements(elements.map(el => selectedIds.includes(el.id) ? { ...el, locked: true } : el));
                                    setSelectedIds([]);
                                }} className="bg-white border hover:bg-gray-50 text-yellow-700 py-1.5 rounded text-[10px] font-bold transition flex items-center justify-center gap-1"><Lock size={14}/> Kunci Semua</button>
                            </div>

                            <div className="pt-2 border-t mt-3">
                                <p className="text-[10px] font-bold text-gray-500 uppercase mb-2">Align (Ratakan)</p>
                                <div className="grid grid-cols-3 gap-1">
                                    <button onClick={() => alignMultiple('left')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center" title="Rata Kiri"><AlignLeft size={16}/></button>
                                    <button onClick={() => alignMultiple('center')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center" title="Rata Tengah Horiz"><AlignCenter size={16}/></button>
                                    <button onClick={() => alignMultiple('right')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center" title="Rata Kanan"><AlignRight size={16}/></button>
                                    <button onClick={() => alignMultiple('top')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center mt-1" title="Rata Atas"><AlignStartVertical size={16}/></button>
                                    <button onClick={() => alignMultiple('middle')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center mt-1" title="Rata Tengah Vertikal"><AlignCenterVertical size={16}/></button>
                                    <button onClick={() => alignMultiple('bottom')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center mt-1" title="Rata Bawah"><AlignEndVertical size={16}/></button>
                                </div>
                            </div>
                            
                            <div className="pt-2 border-t mt-3">
                                <p className="text-[10px] font-bold text-gray-500 uppercase mb-2">Ubah Ukuran Bersama</p>
                                <div className="flex gap-2">
                                    <div className="w-1/2">
                                        <label className="text-[10px] text-gray-500 font-bold uppercase">Set Lebar</label>
                                        <input type="number" className="w-full p-1.5 border rounded text-sm" placeholder="Otomatis" onChange={e => {
                                            const val = Number(e.target.value);
                                            if (val > 0) {
                                                setPast(p => [...p, elements]); setFuture([]);
                                                setElements(elements.map(el => selectedIds.includes(el.id) ? { ...el, width: val } : el));
                                            }
                                        }}/>
                                    </div>
                                    <div className="w-1/2">
                                        <label className="text-[10px] text-gray-500 font-bold uppercase">Set Tinggi</label>
                                        <input type="number" className="w-full p-1.5 border rounded text-sm" placeholder="Otomatis" onChange={e => {
                                            const val = Number(e.target.value);
                                            if (val > 0) {
                                                setPast(p => [...p, elements]); setFuture([]);
                                                setElements(elements.map(el => selectedIds.includes(el.id) ? { ...el, height: val } : el));
                                            }
                                        }}/>
                                    </div>
                                </div>
                            </div>

                            <button onClick={() => {
                                setPast(p => [...p, elements]); setFuture([]);
                                setElements(elements.filter(el => !selectedIds.includes(el.id)));
                                setSelectedIds([]);
                            }} className="w-full mt-3 bg-red-50 hover:bg-red-100 text-red-600 py-2 rounded text-[11px] font-bold flex justify-center items-center gap-1 transition"><Trash2 size={14}/> Hapus Semua Terpilih</button>
                        </div>
                    )}

                    {selectedIds.length === 1 && activeEl && (
                        <div className="space-y-3 pt-2 pb-8 bg-blue-50/30 p-3 rounded-lg border border-blue-100">
                            <p className="text-xs font-bold text-blue-800 uppercase flex items-center justify-between">
                                Sedang Edit: {activeEl.type === 'group' ? 'Grup' : activeEl.type === 'table_grades' ? 'Tabel' : activeEl.type === 'image' ? 'Gambar' : 'Teks'}
                                <button onClick={() => setSelectedIds([])} className="text-gray-400 hover:text-gray-700"><X size={14}/></button>
                            </p>

                            {activeEl.locked ? (
                                <div className="p-4 bg-yellow-50 border border-yellow-200 rounded-lg text-center space-y-3 mt-4">
                                    <Lock size={24} className="mx-auto text-yellow-600"/>
                                    <p className="text-sm font-bold text-yellow-800">Elemen Terkunci</p>
                                    <p className="text-xs text-yellow-700">Buka kunci untuk menggeser atau mengubah isi elemen ini.</p>
                                    <button onClick={() => updateElement(selectedElementId, { locked: false })} className="bg-yellow-500 hover:bg-yellow-600 text-white px-4 py-2 rounded font-bold text-xs transition w-full flex justify-center items-center gap-2"><Lock size={14}/> Buka Kunci</button>
                                </div>
                            ) : (
                                <>
                            {activeEl.type === 'text' && (
                                <textarea className="w-full p-2 border rounded text-sm focus:ring-2 outline-none min-h-[60px]" value={activeEl.content} onChange={e => updateElement(selectedElementId, { content: e.target.value })} />
                            )}

                            {activeEl.type === 'line' && (
                                <div className="space-y-2">
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Warna Garis</label>
                                        <input type="color" value={activeEl.lineColor || '#000000'} onChange={e => updateElement(activeEl.id, { lineColor: e.target.value }, false)} onBlur={e => updateElement(activeEl.id, { lineColor: e.target.value })} className="w-10 h-7 cursor-pointer rounded border" />
                                        <span className="text-xs font-mono">{activeEl.lineColor || '#000000'}</span>
                                    </div>
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Ketebalan (px)</label>
                                        <input type="number" min="1" max="20" value={activeEl.lineThickness || 2} onChange={e => updateElement(activeEl.id, { lineThickness: Number(e.target.value) })} className="w-16 p-1 border rounded text-xs" />
                                    </div>
                                </div>
                            )}
                            {activeEl.type === 'shape' && (
                                <div className="space-y-2">
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Warna Fill</label>
                                        <input type="color" value={activeEl.shapeFill || '#000000'} onChange={e => updateElement(activeEl.id, { shapeFill: e.target.value }, false)} onBlur={e => updateElement(activeEl.id, { shapeFill: e.target.value })} className="w-10 h-7 cursor-pointer rounded border" />
                                    </div>
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Warna Border</label>
                                        <input type="color" value={activeEl.shapeBorderColor || '#000000'} onChange={e => updateElement(activeEl.id, { shapeBorderColor: e.target.value }, false)} onBlur={e => updateElement(activeEl.id, { shapeBorderColor: e.target.value })} className="w-10 h-7 cursor-pointer rounded border" />
                                    </div>
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Tebal Border</label>
                                        <input type="number" min="0" max="20" value={activeEl.shapeBorder || 0} onChange={e => updateElement(activeEl.id, { shapeBorder: Number(e.target.value) })} className="w-16 p-1 border rounded text-xs" />
                                    </div>
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Radius (px)</label>
                                        <input type="number" min="0" max="200" value={activeEl.shapeRadius || 0} onChange={e => updateElement(activeEl.id, { shapeRadius: Number(e.target.value) })} className="w-16 p-1 border rounded text-xs" />
                                    </div>
                                </div>
                            )}
                            {activeEl.type === 'image' && (
                                <div className="bg-white p-2 rounded border space-y-2">
                                    <label className="block text-xs font-semibold text-gray-700">Upload Ulang Gambar:</label>
                                    <input type="file" accept="image/*" className="text-xs w-full" onChange={handleImageUpload} />
                                </div>
                            )}

                            <div className="flex gap-2">
                                <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Posisi X</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={Math.round(activeEl.x || 0)} onChange={e => updateElement(selectedElementId, { x: Number(e.target.value) })}/></div>
                                <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Posisi Y</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={Math.round(activeEl.y || 0)} onChange={e => updateElement(selectedElementId, { y: Number(e.target.value) })}/></div>
                            </div>
                            <div className="flex gap-2 mt-1">
                                <button onClick={() => updateElement(selectedElementId, { x: (canvasWidth - (activeEl.width || 200)) / 2 })} className="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-600 py-1 rounded text-[10px] font-bold transition">Tengah Horiz</button>
                                <button onClick={() => updateElement(selectedElementId, { y: (canvasHeight - (activeEl.height || 30)) / 2 })} className="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-600 py-1 rounded text-[10px] font-bold transition">Tengah Vertikal</button>
                            </div>
                            <div className="flex gap-2 mt-1">
                                <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase" title="Semakin besar angkanya, semakin di atas. Z-Index 0 untuk Watermark.">Lapisan (Z-Index)</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={activeEl.zIndex ?? 1} onChange={e => updateElement(selectedElementId, { zIndex: Number(e.target.value) })}/></div>
                                <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase" title="0 = transparan penuh, 1 = tidak transparan">Transparansi (0-1)</label><input type="number" step="0.1" min="0" max="1" className="w-full p-1.5 border rounded text-sm" value={activeEl.opacity ?? 1} onChange={e => updateElement(selectedElementId, { opacity: Number(e.target.value) })}/></div>
                            </div>

                            {activeEl.type !== 'table_grades' && activeEl.type !== 'image' && activeEl.type !== 'group' && (
                                <>
                                    <div>
                                        <label className="text-[10px] text-gray-500 font-bold uppercase">Jenis Font</label>
                                        <input 
                                            list="font-options" 
                                            className="w-full p-1.5 border rounded text-sm outline-none bg-white" 
                                            value={activeEl.fontFamily || 'Arial, sans-serif'} 
                                            onChange={e => updateElement(selectedElementId, { fontFamily: e.target.value })}
                                            placeholder="Ketik nama font di komputer (contoh: Tahoma)"
                                        />
                                        <datalist id="font-options">
                                            {allFonts.map((font, idx) => <option key={idx} value={font.value}>{font.name}</option>)}
                                        </datalist>
                                    </div>
                                    <div className="flex gap-2">
                                        <div className="w-2/3"><label className="text-[10px] text-gray-500 font-bold uppercase">Ukuran Teks</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={activeEl.fontSize} onChange={e => updateElement(selectedElementId, { fontSize: Number(e.target.value) })}/></div>
                                        <div className="w-1/3 flex items-end"><button onClick={() => updateElement(selectedElementId, { fontWeight: activeEl.fontWeight === 'bold' ? 'normal' : 'bold' })} className={`w-full border p-1.5 rounded text-sm font-bold transition ${activeEl.fontWeight === 'bold' ? 'bg-gray-800 text-white' : 'bg-white hover:bg-gray-100'}`}>B</button></div>
                                    </div>
                                    <div className="mt-2">
                                        <label className="text-[10px] text-gray-500 font-bold uppercase mb-1 block">Rata Teks</label>
                                        <div className="w-full flex items-center border rounded overflow-hidden">
                                            <button onClick={() => updateElement(selectedElementId, { textAlign: 'left' })} className={`flex-1 p-1.5 flex justify-center transition ${(!activeEl.textAlign || activeEl.textAlign === 'left') ? 'bg-gray-800 text-white' : 'bg-white text-gray-600 hover:bg-gray-100'}`} title="Rata Kiri"><AlignLeft size={16}/></button>
                                            <button onClick={() => updateElement(selectedElementId, { textAlign: 'center' })} className={`flex-1 p-1.5 flex justify-center transition border-l ${activeEl.textAlign === 'center' ? 'bg-gray-800 text-white' : 'bg-white text-gray-600 hover:bg-gray-100'}`} title="Rata Tengah"><AlignCenter size={16}/></button>
                                            <button onClick={() => updateElement(selectedElementId, { textAlign: 'right' })} className={`flex-1 p-1.5 flex justify-center transition border-l ${activeEl.textAlign === 'right' ? 'bg-gray-800 text-white' : 'bg-white text-gray-600 hover:bg-gray-100'}`} title="Rata Kanan"><AlignRight size={16}/></button>
                                            <button onClick={() => updateElement(selectedElementId, { textAlign: 'justify' })} className={`flex-1 p-1.5 flex justify-center transition border-l ${activeEl.textAlign === 'justify' ? 'bg-gray-800 text-white' : 'bg-white text-gray-600 hover:bg-gray-100'}`} title="Rata Kiri Kanan (Justify)"><AlignJustify size={16}/></button>
                                        </div>
                                    </div>
                                    <div className="mt-2">
                                        <label className="flex items-center gap-2 text-xs font-semibold text-gray-700 bg-white p-2 border rounded cursor-pointer">
                                            <input type="checkbox" checked={activeEl?.isRtl || false} onChange={e => updateElement(selectedElementId, { isRtl: e.target.checked })} />
                                            Format Teks Arab (RTL)
                                        </label>
                                    </div>
                                </>
                            )}

                            {(activeEl.type === 'image' || activeEl.type === 'table_grades') && (
                                <div className="flex flex-col gap-2">
                                    <div className="flex gap-2">
                                        <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Lebar (Width)</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={activeEl.width} onChange={e => updateElement(selectedElementId, { width: Number(e.target.value) })}/></div>
                                        <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Tinggi (Height)</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={activeEl.height || ''} onChange={e => updateElement(selectedElementId, { height: Number(e.target.value) })}/></div>
                                    </div>
                                    {activeEl.type === 'image' && (
                                        <button onClick={() => updateElement(selectedElementId, { x: 0, y: 0, width: canvasWidth, height: canvasHeight })} className="w-full bg-indigo-50 hover:bg-indigo-100 text-indigo-700 py-1.5 rounded text-[10px] font-bold transition">Paskan ke Ukuran Halaman</button>
                                    )}
                                    {activeEl.type === 'image' && (
                                        <div className="p-2 bg-gray-50 rounded border space-y-2">
                                            <div>
                                                <label className="text-[10px] text-gray-500 font-bold uppercase">Cara Tampil / Crop</label>
                                                <select className="w-full p-1.5 border rounded text-sm outline-none bg-white mt-1" value={activeEl.objectFit || 'contain'} onChange={e => updateElement(selectedElementId, { objectFit: e.target.value })}>
                                                    <option value="contain">Sesuai Asli (Utuh)</option>
                                                    <option value="cover">Potong Kelebihan (Crop)</option>
                                                    <option value="fill">Tarik Penuh (Stretch)</option>
                                                </select>
                                            </div>
                                            {activeEl.objectFit === 'cover' && (
                                                <div>
                                                    <label className="text-[10px] text-gray-500 font-bold uppercase mb-1 block">Geser Posisi Potong</label>
                                                    <div className="flex gap-2">
                                                        <div className="w-1/2 flex flex-col"><span className="text-[9px] text-gray-400">Horizontal</span><input type="range" min="0" max="100" className="w-full" value={activeEl.objectPositionX !== undefined ? activeEl.objectPositionX : 50} onChange={e => updateElement(selectedElementId, { objectPositionX: Number(e.target.value) })} /></div>
                                                        <div className="w-1/2 flex flex-col"><span className="text-[9px] text-gray-400">Vertikal</span><input type="range" min="0" max="100" className="w-full" value={activeEl.objectPositionY !== undefined ? activeEl.objectPositionY : 50} onChange={e => updateElement(selectedElementId, { objectPositionY: Number(e.target.value) })} /></div>
                                                    </div>
                                                </div>
                                            )}
                                        </div>
                                    )}
                                </div>
                            )}

                            {activeEl.type === 'table_grades' && (
                                <div className="mt-4 border-t pt-3 space-y-3">
                                    <div>
                                        <label className="text-[10px] text-gray-500 font-bold uppercase">Filter Kelas</label>
                                        <select
                                            className="w-full p-1.5 border rounded text-sm bg-white outline-none"
                                            value={activeEl.filterClass || ''}
                                            onChange={e => updateElement(selectedElementId, { filterClass: e.target.value })}
                                        >
                                            <option value="">— Semua Kelas —</option>
                                            {Array.from(new Map(classesData.map(c => [normalizeValue(c.name), c])).values()).map(cls => (
                                                <option key={cls.id} value={cls.id}>{cls.name}</option>
                                            ))}
                                        </select>
                                        <p className="text-[10px] text-gray-400 mt-0.5">Pilih kelas agar hanya mata pelajaran kelas tersebut yang tampil.</p>
                                    </div>

                                    <div className="flex items-center justify-between">
                                        <p className="text-xs font-bold text-orange-800">Konfigurasi Kolom Tabel</p>
                                        <button 
                                            onClick={() => {
                                                const newCols = [...(activeEl.columns||[])];
                                                newCols.push({ id: Date.now().toString(), header: 'Kolom Baru', type: 'NILAI', width: 10 });
                                                updateElement(selectedElementId, { columns: newCols });
                                            }}
                                            className="text-[10px] bg-orange-100 text-orange-700 px-2 py-1 rounded font-bold hover:bg-orange-200"
                                        >+ Kolom</button>
                                    </div>
                                    
                                    <label className="flex items-center gap-2 text-xs font-semibold text-gray-700 bg-white p-2 border rounded cursor-pointer">
                                        <input type="checkbox" checked={activeEl?.groupByCategory || false} onChange={e => updateElement(selectedElementId, { groupByCategory: e.target.checked })} />
                                        Kelompokkan per Kategori Pelajaran
                                    </label>
                                    
                                    <label className="flex items-center gap-2 text-xs font-semibold text-gray-700 bg-white p-2 border rounded cursor-pointer mt-1">
                                        <input type="checkbox" checked={activeEl?.isRtl || false} onChange={e => updateElement(selectedElementId, { isRtl: e.target.checked })} />
                                        Format Tabel Arab (RTL & Angka Arab)
                                    </label>
                                    
                                    <label className="flex items-center gap-2 text-xs font-semibold text-gray-700 bg-white p-2 border rounded cursor-pointer mt-1">
                                        <input type="checkbox" checked={activeEl?.isTransparent || false} onChange={e => updateElement(selectedElementId, { isTransparent: e.target.checked })} />
                                        Latar Tabel Transparan (Tanpa Putih)
                                    </label>

                                    <div className="space-y-2 max-h-[300px] overflow-y-auto pr-1 custom-scrollbar mt-2">
                                        {(activeEl.columns || []).map((col, idx) => (
                                            <div key={col.id} className="bg-white p-2 border rounded border-l-4 border-l-orange-400 relative group">
                                                <div className="absolute top-1 right-1 flex gap-0.5 bg-white rounded shadow-sm border p-0.5 z-10">
                                                    <button onClick={() => {
                                                        if (idx === 0) return;
                                                        const newCols = [...activeEl.columns];
                                                        [newCols[idx - 1], newCols[idx]] = [newCols[idx], newCols[idx - 1]];
                                                        updateElement(selectedElementId, { columns: newCols });
                                                    }} disabled={idx === 0} className={`p-0.5 rounded ${idx === 0 ? 'text-gray-300' : 'text-gray-500 hover:text-gray-800 hover:bg-gray-100'}`} title="Geser ke Kiri (Sebelumnya)"><ChevronUp size={14}/></button>
                                                    <button onClick={() => {
                                                        if (idx === activeEl.columns.length - 1) return;
                                                        const newCols = [...activeEl.columns];
                                                        [newCols[idx + 1], newCols[idx]] = [newCols[idx], newCols[idx + 1]];
                                                        updateElement(selectedElementId, { columns: newCols });
                                                    }} disabled={idx === activeEl.columns.length - 1} className={`p-0.5 rounded ${idx === activeEl.columns.length - 1 ? 'text-gray-300' : 'text-gray-500 hover:text-gray-800 hover:bg-gray-100'}`} title="Geser ke Kanan (Selanjutnya)"><ChevronDown size={14}/></button>
                                                    <div className="w-[1px] bg-gray-200 mx-0.5"></div>
                                                    <button onClick={() => {
                                                        const newCols = [...activeEl.columns]; newCols.splice(idx, 1);
                                                        updateElement(selectedElementId, { columns: newCols });
                                                    }} className="p-0.5 text-red-400 hover:text-red-600 hover:bg-red-50 rounded" title="Hapus Kolom"><X size={14}/></button>
                                                </div>
                                                
                                                <input className="w-[85%] text-xs font-bold border-b mb-1 outline-none focus:border-orange-500" value={col.header} onChange={e => {
                                                    const newCols = [...activeEl.columns]; newCols[idx].header = e.target.value;
                                                    updateElement(selectedElementId, { columns: newCols });
                                                }} placeholder="Judul Header"/>
                                                
                                                <div className="flex gap-1 mt-1">
                                                    <select className="w-2/3 text-[10px] p-1 border rounded bg-gray-50" value={col.type} onChange={e => {
                                                        const newCols = [...activeEl.columns]; newCols[idx].type = e.target.value;
                                                        updateElement(selectedElementId, { columns: newCols });
                                                    }}>
                                                        <optgroup label="Standar Pelajaran">
                                                            <option value="NO">Nomor Urut</option>
                                                            <option value="KATEGORI">Kategori Pelajaran</option>
                                                            <option value="MAPEL_ID">Nama Pelajaran (Indo)</option>
                                                            <option value="MAPEL_AR">Nama Pelajaran (Arab)</option>
                                                            <option value="KKM">Nilai KKM</option>
                                                            <option value="NILAI">Nilai Angka Santri</option>
                                                            <option value="RATA_KELAS">Rata-rata Kelas</option>
                                                            <option value="SPASI_KOSONG">Spasi Pemisah (Tanpa Border)</option>
                                                        </optgroup>
                                                        <optgroup label="Pelajaran (Khusus Arab)">
                                                            <option value="NO_AR">Nomor Urut (Arab)</option>
                                                            <option value="KATEGORI_AR">Kategori Pelajaran (Arab)</option>
                                                            <option value="KKM_AR">Nilai KKM (Arab)</option>
                                                            <option value="NILAI_AR">Nilai Angka Santri (Arab)</option>
                                                            <option value="RATA_KELAS_AR">Rata-rata Kelas (Arab)</option>
                                                        </optgroup>
                                                        {data.presences.length > 0 && (
                                                            <optgroup label="Aspek Presensi">
                                                                {data.presences.map(p => <option key={p.id} value={`PRESENCE_${p.id}`}>{p.name}</option>)}
                                                            </optgroup>
                                                        )}
                                                        {data.characterTraits?.length > 0 && (
                                                            <optgroup label="Sikap/Kesantrian">
                                                                {data.characterTraits.map(p => <option key={p.id} value={`SIKAP_${p.id}`}>{p.name}</option>)}
                                                            </optgroup>
                                                        )}
                                                        {data.extracurriculars?.length > 0 && (
                                                            <optgroup label="Ekstrakurikuler">
                                                                {data.extracurriculars.map(p => <option key={p.id} value={`EKSKUL_${p.id}`}>{p.name}</option>)}
                                                            </optgroup>
                                                        )}
                                                    </select>
                                                    <div className="flex w-1/3 gap-1">
                                                        <div className="w-1/2 relative">
                                                            <input type="number" className="w-full text-[10px] p-1 border rounded pr-4" value={col.width} onChange={e => {
                                                                const newCols = [...activeEl.columns]; newCols[idx].width = Number(e.target.value);
                                                                updateElement(selectedElementId, { columns: newCols });
                                                            }} title="Lebar (%)"/>
                                                            <span className="absolute right-1 top-1 text-[10px] text-gray-400 pointer-events-none">%</span>
                                                        </div>
                                                        <div className="w-1/2 relative">
                                                            <input type="number" className="w-full text-[10px] p-1 border rounded pr-4" value={col.height || ''} onChange={e => {
                                                                const newCols = [...activeEl.columns]; newCols[idx].height = Number(e.target.value);
                                                                updateElement(selectedElementId, { columns: newCols });
                                                            }} title="Tinggi Kolom (px)" placeholder="Tinggi"/>
                                                            <span className="absolute right-1 top-1 text-[10px] text-gray-400 pointer-events-none">px</span>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                            )}

                            <div className="flex gap-2 mt-4 flex-wrap">
                                {activeEl.type === 'group' && (
                                    <button onClick={ungroupElements} className="w-full mb-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 py-2 rounded text-[11px] font-bold flex justify-center items-center gap-1 transition"><Layers size={14}/> Ungroup</button>
                                )}
                                <button onClick={() => {updateElement(selectedElementId, { locked: true }); setSelectedIds([]);}} className="flex-1 min-w-[30%] bg-yellow-50 hover:bg-yellow-100 text-yellow-700 py-2 rounded text-[11px] font-bold flex justify-center items-center gap-1 transition" title="Kunci posisi agar tidak tergeser"><Lock size={14}/> Kunci</button>
                                <button onClick={() => duplicateElement(selectedElementId)} className="flex-1 min-w-[30%] bg-blue-50 hover:bg-blue-100 text-blue-600 py-2 rounded text-[11px] font-bold flex justify-center items-center gap-1 transition"><Copy size={14}/> Duplikat</button>
                                <button onClick={() => removeElement(selectedElementId)} className="flex-1 min-w-[30%] bg-red-50 hover:bg-red-100 text-red-600 py-2 rounded text-[11px] font-bold flex justify-center items-center gap-1 transition"><Trash2 size={14}/> Hapus</button>
                            </div>
                                </>
                            )}
                        </div>
                    )}
                </div>
                
                <div className="border-t bg-gray-50 shrink-0 z-10 flex flex-col">
                    <button 
                        onClick={() => setIsBottomMenuExpanded(!isBottomMenuExpanded)} 
                        className="w-full py-1.5 bg-gray-200 hover:bg-gray-300 flex items-center justify-center text-gray-600 text-[10px] font-bold uppercase tracking-wide transition border-b"
                        title={isBottomMenuExpanded ? 'Sembunyikan Opsi Tambahan' : 'Tampilkan Opsi Tambahan'}
                    >
                        {isBottomMenuExpanded ? <ChevronDown size={14} className="mr-1"/> : <ChevronDown size={14} className="mr-1 rotate-180"/>} 
                        {isBottomMenuExpanded ? 'Sembunyikan Opsi' : 'Tampilkan Opsi Menu'}
                    </button>
                    {isBottomMenuExpanded && (
                        <div className="p-4 space-y-2">
                            <p className="text-[10px] text-gray-500 text-center leading-tight">Tarik garis gelap (Atas/Kiri kanvas) untuk Garis Bantu. D-click garis untuk hapus.</p>
                            <button onClick={saveLayout} className="w-full bg-emerald-600 hover:bg-emerald-700 text-white py-3 rounded-lg font-bold flex items-center justify-center gap-2 transition"><Save size={18}/> Simpan Layout</button>
                            <button onClick={() => setShowToolbar(!showToolbar)} className="w-full bg-slate-200 hover:bg-slate-300 text-slate-700 py-2 rounded-lg font-bold flex items-center justify-center gap-2 transition text-sm mt-2">
                                {showToolbar ? <EyeOff size={16}/> : <Eye size={16}/>} {showToolbar ? 'Sembunyikan Menu' : 'Tampilkan Menu'}
                            </button>
                            <button onClick={() => setShowRuler(!showRuler)} className="w-full bg-slate-200 hover:bg-slate-300 text-slate-700 py-2 rounded-lg font-bold flex items-center justify-center gap-2 transition text-sm">
                                {showRuler ? <EyeOff size={16}/> : <Eye size={16}/>} {showRuler ? 'Sembunyikan Ruler' : 'Tampilkan Ruler'}
                            </button>
                            <button onClick={() => setShowGuideBars(!showGuideBars)} className="w-full bg-slate-200 hover:bg-slate-300 text-slate-700 py-2 rounded-lg font-bold flex items-center justify-center gap-2 transition text-sm">
                                {showGuideBars ? <EyeOff size={16}/> : <Eye size={16}/>} {showGuideBars ? 'Sembunyikan Garis Bantu' : 'Tampilkan Garis Bantu'}
                            </button>
                            <button onClick={() => setShowGrid(!showGrid)} className="w-full bg-slate-200 hover:bg-slate-300 text-slate-700 py-2 rounded-lg font-bold flex items-center justify-center gap-2 transition text-sm">
                                {showGrid ? <EyeOff size={16}/> : <Eye size={16}/>} {showGrid ? 'Sembunyikan Grid Latar' : 'Tampilkan Grid Latar'}
                            </button>
                            <button onClick={() => setGuides({ h: [], v: [] })} className="w-full bg-red-100 hover:bg-red-200 text-red-600 py-2 rounded-lg font-bold flex items-center justify-center gap-2 transition text-sm mt-2">
                                <Trash2 size={16}/> Hapus Semua Garis
                            </button>
                        </div>
                    )}
                </div>
            </div>
            )}

            <div id="canvas-scroll-area" onMouseDown={handleCanvasMouseDown} onMouseMove={handleMouseMove} onMouseUp={handleMouseUp} onMouseLeave={handleMouseUp} className="flex-1 bg-gray-200 rounded-xl overflow-auto p-4 flex flex-col items-center border border-gray-300 relative select-none custom-scrollbar print:bg-white print:p-0 print:border-none print:overflow-visible print:static">
                
                {showToolbar && (
                    <div className="bg-white px-4 py-2 rounded-full shadow-sm flex items-center gap-4 mb-4 shrink-0 border border-gray-200 sticky top-0 z-50">
                        <button onClick={() => setShowSidebar(!showSidebar)} className="text-gray-500 hover:text-emerald-600 transition" title={showSidebar ? "Sembunyikan Panel Kiri" : "Tampilkan Panel Kiri"}>
                            <Columns size={18} className={!showSidebar ? "opacity-50" : ""}/>
                        </button>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={() => setCurrentPage(Math.max(0, currentPage - 1))} disabled={currentPage === 0} className="text-gray-500 hover:text-emerald-600 disabled:opacity-30"><ChevronDown className="rotate-90" size={18}/></button>
                        <span className="text-sm font-bold text-gray-700">Halaman {currentPage + 1}</span>
                        <button onClick={() => setCurrentPage(currentPage + 1)} disabled={currentPage >= (elements.length > 0 ? Math.max(...elements.map(e => e.pageIndex || 0)) : 0) + 1} className="text-gray-500 hover:text-emerald-600 disabled:opacity-30"><ChevronDown className="-rotate-90" size={18}/></button>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={() => {
                            const maxPage = elements.length > 0 ? Math.max(...elements.map(e => e.pageIndex || 0)) : 0;
                            setCurrentPage(maxPage + 1);
                        }} className="text-xs font-bold bg-emerald-50 text-emerald-600 px-3 py-1 rounded-full hover:bg-emerald-100 transition mr-2">+ Halaman Baru</button>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={undo} disabled={past.length === 0} className="text-gray-500 hover:text-blue-600 disabled:opacity-30" title="Undo"><Undo size={18}/></button>
                        <button onClick={redo} disabled={future.length === 0} className="text-gray-500 hover:text-blue-600 disabled:opacity-30" title="Redo"><Redo size={18}/></button>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={() => setZoom(z => Math.max(0.3, z - 0.1))} className="text-gray-500 hover:text-emerald-600" title="Zoom Out"><ZoomOut size={18}/></button>
                        <span className="text-sm font-bold text-gray-700 w-10 text-center">{Math.round(zoom * 100)}%</span>
                        <button onClick={() => setZoom(z => Math.min(2, z + 0.1))} className="text-gray-500 hover:text-emerald-600" title="Zoom In"><ZoomIn size={18}/></button>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={toggleFullscreen} className="text-gray-500 hover:text-emerald-600 transition" title={isFullscreen ? "Keluar Fullscreen" : "Layar Penuh (Fullscreen)"}>
                            {isFullscreen ? <Minimize size={18}/> : <Maximize size={18}/>}
                        </button>
                        {selectedIds.length > 0 && (<>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <span className="text-[10px] text-gray-400 font-bold uppercase tracking-wide">Align</span>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('left') : alignElement('left')}   title="Rata Kiri"            className="text-gray-500 hover:text-indigo-600 transition"><AlignLeft size={18}/></button>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('center') : alignElement('center')} title="Tengah Horizontal"   className="text-gray-500 hover:text-indigo-600 transition"><AlignCenter size={18}/></button>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('right') : alignElement('right')}  title="Rata Kanan"           className="text-gray-500 hover:text-indigo-600 transition"><AlignRight size={18}/></button>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('top') : alignElement('top')}    title="Rata Atas"            className="text-gray-500 hover:text-indigo-600 transition"><AlignStartVertical size={18}/></button>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('middle') : alignElement('middle')} title="Tengah Vertikal"     className="text-gray-500 hover:text-indigo-600 transition"><AlignCenterVertical size={18}/></button>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('bottom') : alignElement('bottom')} title="Rata Bawah"           className="text-gray-500 hover:text-indigo-600 transition"><AlignEndVertical size={18}/></button>
                        
                        {selectedIds.length > 1 && (
                            <>
                                <div className="w-px h-4 bg-gray-300 ml-2"></div>
                                <button onClick={groupElements} title="Group Elemen" className="text-gray-500 hover:text-indigo-600 transition ml-2"><Layers size={18}/></button>
                            </>
                        )}
                        {selectedIds.length === 1 && activeEl?.type === 'group' && (
                            <>
                                <div className="w-px h-4 bg-gray-300 ml-2"></div>
                                <button onClick={ungroupElements} title="Ungroup Elemen" className="text-gray-500 hover:text-indigo-600 transition ml-2"><Layers size={18}/></button>
                            </>
                        )}
                        </>)}
                    </div>
                )}

                <div className="relative print:static" style={{ flexShrink: 0, width: `${canvasWidth * zoom + (showRuler ? 40 : 0)}px`, minWidth: `${canvasWidth * zoom + (showRuler ? 40 : 0)}px`, height: `${canvasHeight * zoom + (showRuler ? 28 : 0)}px` }}>
                    {/* Horizontal ruler in cm */}
                    {showRuler && (
                    <div style={{ position: 'absolute', left: 40, top: 0, width: canvasWidth * zoom, height: 28, background: '#334155', display: 'flex', alignItems: 'stretch' }}>
                        <svg width={canvasWidth * zoom} height={28} style={{ display: 'block' }}>
                            <rect x={0} y={0} width={canvasWidth * zoom} height={28} fill="#334155"/>
                            {Array.from({ length: Math.ceil(canvasWidth / PX_PER_CM) + 1 }, (_, i) => {
                                const x = Math.round(i * PX_PER_CM * zoom);
                                if (x > canvasWidth * zoom) return null;
                                return (
                                    <g key={i}>
                                        <line x1={x} y1={18} x2={x} y2={28} stroke="#94a3b8" strokeWidth="1"/>
                                        {i > 0 && <text x={x + 3} y={14} fill="#e2e8f0" fontSize="10" fontFamily="Arial, sans-serif" fontWeight="500">{i}</text>}
                                        {i === 0 && <text x={x + 3} y={14} fill="#94a3b8" fontSize="9" fontFamily="Arial, sans-serif">cm</text>}
                                    </g>
                                );
                            })}
                        </svg>
                    </div>
                    )}
                    {/* Vertical ruler in cm */}
                    {showRuler && (
                    <div style={{ position: 'absolute', left: 0, top: 28, width: 40, height: canvasHeight * zoom, background: '#334155' }}>
                        <svg width={40} height={canvasHeight * zoom} style={{ display: 'block' }}>
                            <rect x={0} y={0} width={40} height={canvasHeight * zoom} fill="#334155"/>
                            {Array.from({ length: Math.ceil(canvasHeight / PX_PER_CM) + 1 }, (_, i) => {
                                const y = Math.round(i * PX_PER_CM * zoom);
                                if (y > canvasHeight * zoom) return null;
                                return (
                                    <g key={i}>
                                        <line x1={30} y1={y} x2={40} y2={y} stroke="#94a3b8" strokeWidth="1"/>
                                        {i > 0 && (
                                            <text
                                                x={20}
                                                y={y - 3}
                                                fill="#e2e8f0"
                                                fontSize="10"
                                                fontFamily="Arial, sans-serif"
                                                fontWeight="500"
                                                textAnchor="middle"
                                                transform={`rotate(-90, 20, ${y - 3})`}
                                            >{i}</text>
                                        )}
                                    </g>
                                );
                            })}
                        </svg>
                    </div>
                    )}
                    {/* Corner box */}
                    {showRuler && <div style={{ position: 'absolute', left: 0, top: 0, width: 40, height: 28, background: '#1e293b' }}/>}
                    {/* Canvas wrapper */}
                    <div className="print-wrapper" style={{ position: 'absolute', left: showRuler ? 40 : 0, top: showRuler ? 28 : 0, width: canvasWidth * zoom, height: canvasHeight * zoom }}>
                        {/* drag bars — only show when ruler is hidden to avoid overlap */}
                        {showGuideBars && !showRuler && (
                            <div onMouseDown={createHGuide} title="Tarik ke bawah untuk buat garis bantu horizontal" className="absolute top-[-22px] left-0 right-0 h-[22px] bg-slate-700 text-slate-300 text-xs flex justify-center items-center cursor-row-resize hover:bg-slate-600 transition select-none rounded-t"><Ruler size={12} className="mr-1"/> Tarik Garis Horizontal</div>
                        )}
                        {showGuideBars && !showRuler && (
                            <div onMouseDown={createVGuide} title="Tarik ke kanan untuk buat garis bantu vertikal" className="absolute left-[-22px] top-0 bottom-0 w-[22px] bg-slate-700 text-slate-300 text-xs flex flex-col justify-center items-center cursor-col-resize hover:bg-slate-600 transition select-none rounded-l" style={{ writingMode: 'vertical-rl', textOrientation: 'mixed' }}>Tarik Vertikal <Ruler size={12} className="mt-1"/></div>
                        )}
                        {/* When ruler IS shown, embed drag triggers inside ruler areas */}
                        {showGuideBars && showRuler && (
                            <div onMouseDown={createHGuide} title="Klik ruler atas untuk buat garis bantu horizontal" className="absolute" style={{ left: 0, top: -28, right: 0, height: 28, cursor: 'row-resize', zIndex: 5 }}/>
                        )}
                        {showGuideBars && showRuler && (
                            <div onMouseDown={createVGuide} title="Klik ruler kiri untuk buat garis bantu vertikal" className="absolute" style={{ left: -40, top: 0, bottom: 0, width: 40, cursor: 'col-resize', zIndex: 5 }}/>
                        )}
                        <div ref={canvasRef} style={{ width: canvasWidth, height: canvasHeight, transform: `scale(${zoom})`, transformOrigin: 'top left', position: 'absolute', top: 0, left: 0 }} className="print-container bg-white shadow-xl">
                        {showGrid && <div className="absolute inset-0 pointer-events-none print:hidden" style={{ backgroundImage: 'linear-gradient(#f0f0f0 1px, transparent 1px), linear-gradient(90deg, #f0f0f0 1px, transparent 1px)', backgroundSize: '20px 20px', opacity: 0.5 }}></div>}
                        
                        {/* Safe Area Visual Guide */}
                        {(margins.top > 0 || margins.bottom > 0 || margins.left > 0 || margins.right > 0) && (
                            <div className="absolute border border-red-400 border-dashed pointer-events-none z-10 print:hidden" style={{ top: `${margins.top}mm`, bottom: `${margins.bottom}mm`, left: `${margins.left}mm`, right: `${margins.right}mm` }}>
                                <div className="absolute -top-4 -left-[1px] text-[9px] text-red-500 font-bold bg-white px-1">Batas Area Aman Printer</div>
                            </div>
                        )}

                        {guides.v.map((gx, i) => (<div key={`v-${i}`} onMouseDown={(e) => startDragGuide(e, 'v', i)} onDoubleClick={() => setGuides(prev => ({...prev, v: prev.v.filter((_, idx) => idx !== i)}))} style={{ position: 'absolute', left: `${gx}px`, top: 0, bottom: 0, borderLeft: '1px dashed #0ea5e9', cursor: 'col-resize', zIndex: 10 }} className="hover:border-l-2 hover:border-blue-500 group print:hidden"><div className="absolute -top-5 -left-4 bg-blue-500 text-white text-[10px] px-1 rounded opacity-0 group-hover:opacity-100">{Math.round(gx)}</div></div>))}
                        {guides.h.map((gy, i) => (<div key={`h-${i}`} onMouseDown={(e) => startDragGuide(e, 'h', i)} onDoubleClick={() => setGuides(prev => ({...prev, h: prev.h.filter((_, idx) => idx !== i)}))} style={{ position: 'absolute', top: `${gy}px`, left: 0, right: 0, borderTop: '1px dashed #0ea5e9', cursor: 'row-resize', zIndex: 10 }} className="hover:border-t-2 hover:border-blue-500 group print:hidden"><div className="absolute -left-6 -top-2 bg-blue-500 text-white text-[10px] px-1 rounded opacity-0 group-hover:opacity-100">{Math.round(gy)}</div></div>))}

                        {elements.filter(el => (el.pageIndex || 0) === currentPage).map(el => {
                            const isSelected = selectedIds.includes(el.id);
                            const isDraggingThis = (draggingType === 'element' && dragIndex === el.id) || (draggingType === 'multi_element' && selectedIds.includes(el.id));
                            
                            return (
                                <div key={el.id} data-element-id={el.id} onMouseDown={(e) => handleElementMouseDown(e, el)}
                                    style={{
                                        position: 'absolute', left: `${el.x}px`, top: `${el.y}px`, fontSize: `${el.fontSize}px`, fontFamily: el.fontFamily || 'Arial, sans-serif', fontWeight: el.fontWeight,
                                        width: el.width ? `${el.width}px` : 'auto', height: (el.type === 'image' || el.type === 'table_grades') ? (el.height ? `${el.height}px` : 'auto') : 'auto',
                                        cursor: isDraggingThis ? 'grabbing' : 'grab', outline: isSelected ? '2px dashed #059669' : 'none', padding: (el.type === 'image' || el.type === 'table_grades') ? '0' : '2px',
                                        zIndex: isSelected ? 20 : (el.zIndex ?? 1), opacity: el.opacity ?? 1,
                                        pointerEvents: el.locked ? 'none' : 'auto'
                                    }}
                                    className={`hover:outline hover:outline-1 hover:outline-gray-400 ${el.type === 'table_grades' && !el.isTransparent ? 'bg-white' : ''}`}
                                >
                                    {el.type === 'group' ? (
                                        <div style={{ position: 'relative', width: '100%', height: '100%' }}>
                                            {(el.children || []).map(child => (
                                                <div key={child.id} style={{
                                                    position: 'absolute', left: `${child.x}px`, top: `${child.y}px`, fontSize: `${child.fontSize}px`, fontFamily: child.fontFamily || 'Arial, sans-serif', fontWeight: child.fontWeight,
                                                    width: child.width ? `${child.width}px` : 'auto', height: child.type === 'image' ? `${child.height}px` : 'auto',
                                                    padding: (child.type === 'image' || child.type === 'table_grades') ? '0' : '2px',
                                                    width: child.width ? `${child.width}px` : 'auto', height: (child.type === 'image' || child.type === 'shape') ? `${child.height}px` : 'auto',
                                                    padding: (child.type === 'image' || child.type === 'table_grades' || child.type === 'shape' || child.type === 'line') ? '0' : '2px',
                                                    zIndex: child.zIndex ?? 1, opacity: child.opacity ?? 1
                                                }}>
                                                    {child.type === 'table_grades' ? renderDynamicTable(child, data, mockStudentGrades, mockClassAverages, false, 'raport', classesData) 
                                                    : child.type === 'image' ? <img src={child.content} style={{ width: '100%', height: '100%', objectFit: child.objectFit || 'contain', objectPosition: `${child.objectPositionX ?? 50}% ${child.objectPositionY ?? 50}%`, pointerEvents: 'none' }} alt="elemen" />
                                                    : child.type === 'line' ? <div style={{ width: '100%', height: `${child.lineThickness || 2}px`, backgroundColor: child.lineColor || '#000000', pointerEvents: 'none' }} />
                                                    : child.type === 'shape' ? <div style={{ width: '100%', height: '100%', backgroundColor: child.shapeFill || '#000000', borderRadius: `${child.shapeRadius || 0}px`, border: child.shapeBorder ? `${child.shapeBorder}px solid ${child.shapeBorderColor || '#000000'}` : 'none', pointerEvents: 'none' }} />
                                                    : <div style={{ whiteSpace: 'pre-wrap', width: '100%', height: '100%', textAlign: child.textAlign || 'left', direction: child.isRtl ? 'rtl' : 'ltr' }}>{child.content}</div>}
                                                </div>
                                            ))}
                                        </div>
                                    ) : el.type === 'table_grades' ? renderDynamicTable(el, data, mockStudentGrades, mockClassAverages, false, 'raport', classesData) 
                                    : el.type === 'image' ? <img src={el.content} style={{ width: '100%', height: '100%', objectFit: el.objectFit || 'contain', objectPosition: `${el.objectPositionX ?? 50}% ${el.objectPositionY ?? 50}%`, pointerEvents: 'none' }} alt="elemen" />
                                    : el.type === 'line' ? <div style={{ width: '100%', height: `${el.lineThickness || 2}px`, backgroundColor: el.lineColor || '#000000', pointerEvents: 'none' }} />
                                    : el.type === 'shape' ? <div style={{ width: '100%', height: '100%', backgroundColor: el.shapeFill || '#000000', borderRadius: `${el.shapeRadius || 0}px`, border: el.shapeBorder ? `${el.shapeBorder}px solid ${el.shapeBorderColor || '#000000'}` : 'none', pointerEvents: 'none' }} />
                                    : <div style={{ whiteSpace: 'pre-wrap', width: '100%', height: '100%', textAlign: el.textAlign || 'left', direction: el.isRtl ? 'rtl' : 'ltr' }}>{el.content}</div>}
                                    
                                    {isSelected && !el.locked && selectedIds.length === 1 && (
                                        <>
                                            <div className="absolute -top-3 -left-3 bg-emerald-600 text-white rounded-full p-1 shadow z-30 cursor-move" onMouseDown={(e) => handleElementMouseDown(e, el)}><GripHorizontal size={12} /></div>
                                            <div className="absolute top-[-4px] left-1/2 w-2 h-2 bg-emerald-600 border border-white cursor-n-resize" style={{transform: 'translateX(-50%)'}} onMouseDown={(e) => handleResizeMouseDown(e, el, 'n')}/>
                                            <div className="absolute bottom-[-4px] left-1/2 w-2 h-2 bg-emerald-600 border border-white cursor-s-resize" style={{transform: 'translateX(-50%)'}} onMouseDown={(e) => handleResizeMouseDown(e, el, 's')}/>
                                            <div className="absolute left-[-4px] top-1/2 w-2 h-2 bg-emerald-600 border border-white cursor-w-resize" style={{transform: 'translateY(-50%)'}} onMouseDown={(e) => handleResizeMouseDown(e, el, 'w')}/>
                                            <div className="absolute right-[-4px] top-1/2 w-2 h-2 bg-emerald-600 border border-white cursor-e-resize" style={{transform: 'translateY(-50%)'}} onMouseDown={(e) => handleResizeMouseDown(e, el, 'e')}/>
                                            <div className="absolute top-[-4px] left-[-4px] w-2 h-2 bg-emerald-600 border border-white cursor-nw-resize" onMouseDown={(e) => handleResizeMouseDown(e, el, 'nw')}/>
                                            <div className="absolute top-[-4px] right-[-4px] w-2 h-2 bg-emerald-600 border border-white cursor-ne-resize" onMouseDown={(e) => handleResizeMouseDown(e, el, 'ne')}/>
                                            <div className="absolute bottom-[-4px] left-[-4px] w-2 h-2 bg-emerald-600 border border-white cursor-sw-resize" onMouseDown={(e) => handleResizeMouseDown(e, el, 'sw')}/>
                                            <div className="absolute bottom-[-4px] right-[-4px] w-2 h-2 bg-emerald-600 border border-white cursor-se-resize" onMouseDown={(e) => handleResizeMouseDown(e, el, 'se')}/>
                                        </>
                                    )}
                                </div>
                            );
                        })}

                        {draggingType === 'selection' && selectionBox && (
                            <div style={{
                                position: 'absolute',
                                left: Math.min(selectionBox.startX, selectionBox.endX),
                                top: Math.min(selectionBox.startY, selectionBox.endY),
                                width: Math.abs(selectionBox.endX - selectionBox.startX),
                                height: Math.abs(selectionBox.endY - selectionBox.startY),
                                backgroundColor: 'rgba(59, 130, 246, 0.2)',
                                border: '1px solid rgb(59, 130, 246)',
                                pointerEvents: 'none',
                                zIndex: 9999
                            }} />
                        )}
                    </div>{/* end canvasRef */}
                    </div>{/* end canvas position wrapper */}
                </div>{/* end ruler+canvas outer */}
            </div>
            
            <style>{`
            @media print { 
                body * { visibility: hidden; } 
                .print-wrapper, .print-wrapper * { visibility: visible; } 
                .print-wrapper { position: absolute !important; left: 0 !important; top: 0 !important; width: 100% !important; margin: 0 !important; padding: 0 !important; } 
                .print-container { position: relative !important; margin: 0 !important; padding: 0 !important; box-shadow: none !important; border: none !important; transform: scale(1) !important; left: 0 !important; top: 0 !important; } 
                @page { size: ${pageSize === 'F4' ? '215.9mm 330.2mm' : 'A4'} ${orientation}; margin: 0 !important; } 
            }
            `}</style>
            <style>{`.custom-scrollbar::-webkit-scrollbar { width: 6px; } .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }`}</style>
        </div>
    );
};

// ==========================================
// ==========================================
// UTILITY: EXCEL IMPORT/EXPORT FOR GRADES
// ==========================================
const exportGradesToExcel = (grades, studentsInClass, subjectsInClass, className, activeInputTab, data) => {
    let headers = ['No', 'NIS', 'Nama Santri'];
    let cols = [];
    
    if (activeInputTab === 'pelajaran') {
        subjectsInClass.forEach(sub => {
            headers.push(`${sub.nameId} - UTS`);
            headers.push(`${sub.nameId} - UAS`);
            cols.push(sub.id);
        });
    } else if (activeInputTab === 'presensi') {
        data.presences.forEach(p => { headers.push(p.name); cols.push(p.id); });
    } else if (activeInputTab === 'sikap') {
        data.characterTraits.forEach(p => { headers.push(p.name); cols.push(p.id); });
    } else if (activeInputTab === 'ekskul') {
        data.extracurriculars.forEach(p => { headers.push(p.name); cols.push(p.id); });
    } else if (activeInputTab === 'catatan_wali') {
        headers.push('Catatan Wali Kelas');
        cols.push('catatan_wali');
    }

    const rows = [headers];
    studentsInClass.forEach((st, idx) => {
        const row = [idx + 1, st.nis || '', st.nama];
        if (activeInputTab === 'pelajaran') {
            subjectsInClass.forEach(sub => {
                row.push(grades[st.id]?.[sub.id]?.uts || '');
                row.push(grades[st.id]?.[sub.id]?.uas || '');
            });
        } else if (['presensi', 'sikap', 'ekskul'].includes(activeInputTab)) {
            cols.forEach(colId => {
                row.push(grades[st.id]?.[colId] || '');
            });
        } else if (activeInputTab === 'catatan_wali') {
            row.push(grades[st.id]?.['catatan_wali'] || '');
        }
        rows.push(row);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const colWidths = [8, 15, 25];
    cols.forEach(() => {
        if (activeInputTab === 'pelajaran') { colWidths.push(12, 12); }
        else if (activeInputTab === 'catatan_wali') { colWidths.push(50); }
        else { colWidths.push(15); }
    });
    ws['!cols'] = colWidths.map(w => ({ wch: w }));
    ws['!freeze'] = { xSplit: 3, ySplit: 1 };
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Data_${activeInputTab}`);
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `template_${activeInputTab}_${className.replace(/\//g, '-')}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
};

const importGradesFromExcel = async (file, studentsInClass, subjectsInClass, activeInputTab, data) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const dataBuf = e.target.result;
                const workbook = XLSX.read(dataBuf, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                
                const importedGrades = {};
                rows.forEach((row, rowIdx) => {
                    const nisKey = Object.keys(row).find(k => k.toLowerCase().includes('nis'));
                    const namaKey = Object.keys(row).find(k => k.toLowerCase().includes('nama') && !k.toLowerCase().includes('arab'));
                    
                    const nis = row[nisKey] || '';
                    const nama = row[namaKey] || '';
                    
                    const student = studentsInClass.find(st => 
                        (st.nis && String(st.nis) === String(nis)) ||
                        (st.nama && st.nama.toLowerCase().includes(String(nama).toLowerCase()))
                    );
                    
                    if (!student) {
                        console.warn(`Siswa dengan NIS ${nis} atau nama ${nama} tidak ditemukan.`);
                        return;
                    }
                    
                    if (!importedGrades[student.id]) {
                        importedGrades[student.id] = {};
                    }
                    
                    if (activeInputTab === 'pelajaran') {
                        subjectsInClass.forEach(sub => {
                            const utsKey = Object.keys(row).find(k => k.includes(sub.nameId) && k.includes('UTS'));
                            const uasKey = Object.keys(row).find(k => k.includes(sub.nameId) && k.includes('UAS'));
                            
                            const uts = utsKey ? String(row[utsKey]).trim() : '';
                            const uas = uasKey ? String(row[uasKey]).trim() : '';
                            
                            if (uts || uas) {
                                importedGrades[student.id][sub.id] = {
                                    uts: uts ? convertArabicToLatin(uts) : '',
                                    uas: uas ? convertArabicToLatin(uas) : ''
                                };
                            }
                        });
                    } else if (activeInputTab === 'presensi') {
                        data.presences.forEach(p => {
                            const valKey = Object.keys(row).find(k => k.toLowerCase().includes(p.name.toLowerCase()));
                            if (valKey && row[valKey]) importedGrades[student.id][p.id] = convertArabicToLatin(String(row[valKey]).trim());
                        });
                    } else if (activeInputTab === 'sikap') {
                        data.characterTraits.forEach(p => {
                            const valKey = Object.keys(row).find(k => k.toLowerCase().includes(p.name.toLowerCase()));
                            if (valKey && row[valKey]) importedGrades[student.id][p.id] = String(row[valKey]).trim();
                        });
                    } else if (activeInputTab === 'ekskul') {
                        data.extracurriculars.forEach(p => {
                            const valKey = Object.keys(row).find(k => k.toLowerCase().includes(p.name.toLowerCase()));
                            if (valKey && row[valKey]) importedGrades[student.id][p.id] = String(row[valKey]).trim();
                        });
                    } else if (activeInputTab === 'catatan_wali') {
                        const valKey = Object.keys(row).find(k => k.toLowerCase().includes('catatan wali'));
                        if (valKey && row[valKey]) importedGrades[student.id]['catatan_wali'] = String(row[valKey]).trim();
                    }
                });
                
                resolve(importedGrades);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error('Gagal membaca file'));
        reader.readAsArrayBuffer(file);
    });
};

// INPUT NILAI 
// ==========================================
const InputNilai = ({ activeInputTab }) => {
    const { data, allData, saveToDb, showNotification } = useContext(AppContext);
    const [selectedClass, setSelectedClass] = useState('');
    const [localGrades, setLocalGrades] = useState({});
    const [isSaving, setIsSaving] = useState(false);
    const [lastSaved, setLastSaved] = useState(null);
    const [isInitialized, setIsInitialized] = useState(false);
    const [isImporting, setIsImporting] = useState(false);
    const [zoomLevel, setZoomLevel] = useState(100);
    const [isHeaderHidden, setIsHeaderHidden] = useState(false);
    const [isFullscreen, setIsFullscreen] = useState(false);
    const containerRef = useRef(null);
    
    // Gunakan useRef untuk melacak status terakhir yang disave ke database (Debouncing Check)
    const lastSavedGradesRef = useRef(null);

    const activeSetting = data.settings.find(s => s.isActive);
    const activeStudents = getStudentsForYear(data.studentSnapshots, activeSetting, data.students);
    const classesData = allData?.classes || data.classes;
    const studentsInClass = getStudentsInClass(activeStudents, classesData, selectedClass);
    const subjectsInClass = useMemo(() => sortSubjectsByCategory(filterSubjectsByClass(data.subjects, selectedClass, classesData), data.subjectCategories), [data.subjects, data.subjectCategories, selectedClass, classesData]);
    const gradeDocId = getGradeDocId(selectedClass, classesData, activeSetting, data.grades);

    useEffect(() => {
        if (!gradeDocId) {
            setLocalGrades({}); setIsInitialized(false); return;
        }
        const classGrades = data.grades.find(g => g.id === gradeDocId)?.data || {};
        const initialGrades = {};
        studentsInClass.forEach(st => {
            const raw = classGrades[st.id] || {};
            const normalized = {};
            Object.entries(raw).forEach(([key, value]) => {
                const isSubject = data.subjects.some(sub => sub.id === key);
                if (isSubject) {
                    if (value && typeof value === 'object' && ('uts' in value || 'uas' in value)) {
                        normalized[key] = { uts: value.uts || '', uas: value.uas || '', ...value };
                    } else {
                        normalized[key] = { uts: value == null ? '' : String(value), uas: '' };
                    }
                } else {
                    normalized[key] = value;
                }
            });
            initialGrades[st.id] = normalized;
        });
        
        setLocalGrades(initialGrades);
        // Tandai initialGrades ini sebagai acuan versi data yang terakhir 'disimpan'
        lastSavedGradesRef.current = JSON.stringify(initialGrades);
        setIsInitialized(true);
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [selectedClass, gradeDocId]); 

    const handleGradeChange = (studentId, fieldId, val, category = null) => {
        // Konversi angka Arab ke Latin terjadi di global event listener
        setLocalGrades(prev => {
            const studentGrades = prev[studentId] || {};
            if (category === 'uts' || category === 'uas') {
                const existing = studentGrades[fieldId] && typeof studentGrades[fieldId] === 'object'
                    ? { ...studentGrades[fieldId] }
                    : { uts: '', uas: '' };
                existing[category] = val;
                return { ...prev, [studentId]: { ...studentGrades, [fieldId]: existing } };
            }
            return { ...prev, [studentId]: { ...studentGrades, [fieldId]: val } };
        });
    };

    useEffect(() => {
        if (!isInitialized || !gradeDocId) return;

        const currentGradesStr = JSON.stringify(localGrades);
        
        // CEK KOTOR (DIRTY CHECK): Jangan jalankan fungsi simpan jika data ketikan tidak berubah dari yang terakhir di-save.
        // Ini yang akan mencegah aplikasi Anda auto-save terus menerus tiada henti!
        if (lastSavedGradesRef.current === currentGradesStr) return;

        const timer = setTimeout(async () => {
            setIsSaving(true);
            await saveToDb('grades', gradeDocId, { data: localGrades, class: selectedClass, tahun: activeSetting.tahun, semester: activeSetting.semester }, true);
            lastSavedGradesRef.current = currentGradesStr; // Perbarui acuan dengan data yang baru saja disave
            setIsSaving(false); setLastSaved(new Date());
        }, 1500); 
        return () => clearTimeout(timer);
    }, [localGrades, isInitialized, gradeDocId, activeSetting, saveToDb, selectedClass]);

    const handleManualSave = async () => {
        if (!gradeDocId) return;
        setIsSaving(true);
        await saveToDb('grades', gradeDocId, { data: localGrades, class: selectedClass, tahun: activeSetting.tahun, semester: activeSetting.semester }, false, `Menyimpan data Input Nilai kelas ${selectedClass}`);
        lastSavedGradesRef.current = JSON.stringify(localGrades);
        setIsSaving(false); setLastSaved(new Date());
    };

    const handleExportGrades = () => {
        if (!selectedClass || (activeInputTab === 'pelajaran' && subjectsInClass.length === 0)) {
            showNotification('Pilih kelas dan pastikan ada mata pelajaran.', 'error');
            return;
        }
        const className = getClassNameFromValue(classesData, selectedClass);
        exportGradesToExcel(localGrades, studentsInClass, subjectsInClass, className, activeInputTab, data);
    };

    const handleImportGrades = async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        
        if (!selectedClass || (activeInputTab === 'pelajaran' && subjectsInClass.length === 0)) {
            showNotification('Pilih kelas dan pastikan ada mata pelajaran.', 'error');
            return;
        }

        try {
            setIsImporting(true);
            const importedGrades = await importGradesFromExcel(file, studentsInClass, subjectsInClass, activeInputTab, data);
            
            // Merge dengan grades yang sudah ada
            setLocalGrades(prev => {
                const merged = { ...prev };
                Object.entries(importedGrades).forEach(([studentId, subjectGrades]) => {
                    merged[studentId] = { ...prev[studentId], ...subjectGrades };
                });
                return merged;
            });
            
            showNotification(`✓ ${Object.keys(importedGrades).length} siswa berhasil diimpor!`);
            // Reset file input
            e.target.value = '';
        } catch (err) {
            console.error(err);
            showNotification('Gagal mengimpor file Excel. Pastikan format sesuai template.', 'error');
        } finally {
            setIsImporting(false);
        }
    };

    const handleZoom = (direction) => {
        setZoomLevel(prev => {
            const newLevel = direction === 'in' ? prev + 10 : prev - 10;
            return Math.max(50, Math.min(200, newLevel));
        });
    };

    const handleResetZoom = () => {
        setZoomLevel(100);
    };

    const toggleHeaderHidden = () => {
        setIsHeaderHidden(!isHeaderHidden);
    };

    const toggleFullscreen = async () => {
        if (!containerRef.current) return;
        
        try {
            if (!document.fullscreenElement) {
                await containerRef.current.requestFullscreen();
                setIsFullscreen(true);
            } else {
                await document.exitFullscreen();
                setIsFullscreen(false);
            }
        } catch (err) {
            console.error('Fullscreen error:', err);
            showNotification('Fullscreen tidak tersedia di browser Anda', 'error');
        }
    };

    // Listen for fullscreen change
    useEffect(() => {
        const handleFullscreenChange = () => {
            setIsFullscreen(!!document.fullscreenElement);
        };
        
        document.addEventListener('fullscreenchange', handleFullscreenChange);
        return () => document.removeEventListener('fullscreenchange', handleFullscreenChange);
    }, []);

    if (!activeSetting) return (
        <div className="bg-white rounded-xl shadow-sm p-6 border border-gray-100 text-center py-12"><AlertCircle className="mx-auto text-yellow-500 mb-4" size={48} /><h3 className="text-xl font-bold text-gray-800 mb-2">Tahun Ajaran Belum Aktif</h3><p className="text-gray-500">Silakan Minta Admin mengaktifkan Tahun Ajaran di Master Data.</p></div>
    );

    const classTotals = {}; const classCounts = {};
    subjectsInClass.forEach(sub => { classTotals[sub.id] = 0; classCounts[sub.id] = 0; });

    const renderTableContent = () => {
        if (activeInputTab === 'pelajaran' && subjectsInClass.length === 0) {
            return (
                <div className="bg-white rounded-xl shadow-sm p-6 border border-gray-100 text-center">
                    <h3 className="text-lg font-semibold text-gray-800">Belum ada mata pelajaran untuk kelas ini</h3>
                    <p className="text-gray-500">Atur pelajaran di Master Data dan pilih kelas yang sesuai untuk menampilkan Input Nilai.</p>
                </div>
            );
        }
        if (activeInputTab === 'pelajaran') {
            const groupedSubjects = groupBy(subjectsInClass, 'kategori');
            const orderedGroups = Object.entries(groupedSubjects).sort(([aKey], [bKey]) => {
                const aIs = isReligiousCategory(aKey);
                const bIs = isReligiousCategory(bKey);
                if (aIs !== bIs) return aIs ? -1 : 1;
                return (aKey || '').localeCompare(bKey || '');
            });
            return (
                <table className="w-full text-left border-collapse whitespace-nowrap">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-emerald-700 text-white text-sm">
                            <th rowSpan={3} className="p-3 border-b border-r border-emerald-600 text-center w-12 sticky left-0 z-30 bg-emerald-800">No</th>
                            <th rowSpan={3} className="p-3 border-b border-r border-emerald-600 text-center w-20 bg-emerald-800">NIS</th>
                            <th rowSpan={3} className="p-3 border-b border-r border-emerald-600 sticky left-12 z-30 bg-emerald-800">Nama Santri</th>
                            {subjectsInClass.length > 0 && <th colSpan={subjectsInClass.length} className="p-2 border-b border-r border-emerald-600 text-center font-bold bg-emerald-800">NILAI MATA PELAJARAN</th>}
                            <th rowSpan={3} className="p-3 border-b border-r border-emerald-600 text-center w-20 bg-emerald-900">Total<br/>Raport</th>
                            <th rowSpan={3} className="p-3 border-b border-emerald-600 text-center w-20 bg-emerald-900">Rata-rata<br/>Raport</th>
                        </tr>
                        {subjectsInClass.length > 0 && (
                        <tr className="bg-emerald-600 text-white text-sm">
                            {orderedGroups.map(([cat, subs]) => (
                                <th key={cat || 'umum'} colSpan={subs.length} className="p-2 border-b border-r border-emerald-500 text-center bg-emerald-700">
                                    {cat || 'Umum'}
                                </th>
                            ))}
                        </tr>
                        )}
                        <tr className="bg-emerald-600 text-white text-sm">
                            {subjectsInClass.map(sub => (
                                <th key={sub.id} className="p-3 border-b border-r border-emerald-500 text-center min-w-[210px] bg-emerald-700">
                                    <div className="font-bold">{sub.nameId}</div>
                                    <div className="text-[11px] text-emerald-200 font-normal mt-0.5">(Guru: {sub.guru || '-'})</div>
                                    <div className="text-[11px] text-yellow-300 font-bold mt-0.5">KKM: {sub.kkm || '-'}</div>
                                    <div className="text-[11px] text-slate-100 font-semibold mt-2">UTS / UAS / Raport</div>
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {studentsInClass.map((st, idx) => {
                            let rowRaportTotal = 0; let rowRaportCount = 0;
                            return (
                                <tr key={st.id} className="border-b hover:bg-gray-50 transition-colors">
                                    <td className="p-3 text-center text-gray-500 sticky left-0 bg-white border-r z-10">{idx + 1}</td>
                                    <td className="p-3 text-center bg-white border-r font-semibold">{st.nis || '-'}</td>
                                    <td className="p-3 font-semibold sticky left-12 bg-white border-r shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)] z-10 text-gray-800">{st.nama}</td>
                                    {subjectsInClass.map(sub => {
                                        const uts = localGrades[st.id]?.[sub.id]?.uts || '';
                                        const uas = localGrades[st.id]?.[sub.id]?.uas || '';
                                        const raport = computeRaportScore(uts, uas);
                                        if (raport !== '') { rowRaportTotal += raport; rowRaportCount++; classTotals[sub.id] += raport; classCounts[sub.id]++; }
                                        const isRed = raport !== '' && raport < Number(sub.kkm || 0);
                                        return (
                                        <td key={sub.id} className="p-2 border-r bg-white hover:bg-emerald-50">
                                                <div className="grid grid-cols-3 gap-1 items-center">
                                                    <input type="text" dir="auto" title="Nilai UTS (angka)" placeholder="UTS" className="w-full min-w-[56px] p-2 border rounded text-center text-base font-bold outline-none text-gray-800 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-300" value={uts} onChange={e => handleGradeChange(st.id, sub.id, e.target.value, 'uts')} />
                                                    <input type="text" dir="auto" title="Nilai UAS (angka)" placeholder="UAS" className="w-full min-w-[56px] p-2 border rounded text-center text-base font-bold outline-none text-gray-800 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-300" value={uas} onChange={e => handleGradeChange(st.id, sub.id, e.target.value, 'uas')} />
                                                    <div className={`rounded border p-2 text-base font-bold text-center min-w-[56px] ${isRed ? 'text-red-600 bg-red-50 border-red-200' : 'text-gray-800 bg-gray-50 border-gray-200'}`}>
                                                        {raport === '' ? '-' : raport}
                                                    </div>
                                                </div>
                                            </td>
                                        );
                                    })}
                                    <td className="p-3 text-center font-bold text-emerald-800 bg-emerald-50 border-r">{rowRaportTotal !== 0 ? rowRaportTotal.toFixed(2) : '-'}</td>
                                    <td className="p-3 text-center font-bold text-blue-800 bg-blue-50">{rowRaportCount > 0 ? (rowRaportTotal / rowRaportCount).toFixed(2) : '-'}</td>
                                </tr>
                            );
                        })}
                    </tbody>
                    <tfoot className="sticky bottom-0 z-20 shadow-[0_-2px_5px_rgba(0,0,0,0.05)]">
                        <tr className="bg-gray-100 text-gray-800">
                            <td colSpan="3" className="p-3 text-right font-bold border-r sticky left-0 z-30 bg-gray-200">Rata-rata Raport per Pelajaran</td>
                            {subjectsInClass.map(sub => (
                                <td key={sub.id} className="p-3 text-center font-bold border-r text-blue-700">{classCounts[sub.id] > 0 ? (classTotals[sub.id] / classCounts[sub.id]).toFixed(2) : '-'}</td>
                            ))}
                            <td className="p-3 text-center font-bold border-r text-blue-700">-</td>
                            <td className="bg-gray-200 border-l"></td>
                        </tr>
                    </tfoot>
                </table>
            );
        }
        
        if (activeInputTab === 'presensi') {
            return (
                <table className="w-full text-left border-collapse whitespace-nowrap">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-indigo-700 text-white text-sm">
                            <th className="p-3 border-b border-r border-indigo-600 text-center w-12 sticky left-0 z-30 bg-indigo-800">No</th>
                            <th className="p-3 border-b border-r border-indigo-600 text-center w-20 bg-indigo-800">NIS</th>
                            <th className="p-3 border-b border-r border-indigo-600 sticky left-12 z-30 bg-indigo-800">Nama Santri</th>
                            {data.presences.map(p => <th key={p.id} className="p-3 border-b border-r border-indigo-600 text-center min-w-[100px]">{p.name}</th>)}
                        </tr>
                    </thead>
                    <tbody>
                        {studentsInClass.map((st, idx) => (
                            <tr key={st.id} className="border-b hover:bg-gray-50 transition-colors">
                                <td className="p-3 text-center text-gray-500 sticky left-0 bg-white border-r z-10">{idx + 1}</td>
                                <td className="p-3 text-center bg-white border-r font-semibold">{st.nis || '-'}</td>
                                <td className="p-3 font-semibold sticky left-12 bg-white border-r shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)] z-10">{st.nama}</td>
                                {data.presences.map(p => (
                                    <td key={p.id} className="p-2 border-r bg-white hover:bg-indigo-50">
                                        <input type="text" dir="auto" title="Ketik angka Arab atau Latin (٠-٩ atau 0-9)" className="w-full p-2 border rounded text-center font-bold outline-none focus:border-indigo-500 text-indigo-900"
                                            value={localGrades[st.id]?.[p.id] || ''} onChange={e => handleGradeChange(st.id, p.id, e.target.value)} placeholder="-" />
                                    </td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
            );
        }

        if (activeInputTab === 'sikap') {
            return (
                <table className="w-full text-left border-collapse whitespace-nowrap">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-blue-700 text-white text-sm">
                            <th className="p-3 border-b border-r border-blue-600 text-center w-12 sticky left-0 z-30 bg-blue-800">No</th>
                            <th className="p-3 border-b border-r border-blue-600 text-center w-20 bg-blue-800">NIS</th>
                            <th className="p-3 border-b border-r border-blue-600 sticky left-12 z-30 bg-blue-800">Nama Santri</th>
                            {data.characterTraits.map(p => <th key={p.id} className="p-3 border-b border-r border-blue-600 text-center min-w-[120px]">{p.name}</th>)}
                        </tr>
                    </thead>
                    <tbody>
                        {studentsInClass.map((st, idx) => (
                            <tr key={st.id} className="border-b hover:bg-gray-50 transition-colors">
                                <td className="p-3 text-center text-gray-500 sticky left-0 bg-white border-r z-10">{idx + 1}</td>
                                <td className="p-3 text-center bg-white border-r font-semibold">{st.nis || '-'}</td>
                                <td className="p-3 font-semibold sticky left-12 bg-white border-r shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)] z-10">{st.nama}</td>
                                {data.characterTraits.map(p => (
                                    <td key={p.id} className="p-2 border-r bg-white hover:bg-blue-50">
                                        <input type="text" dir="auto" title="Ketik nilai (A/B/C atau angka Arab/Latin)" className="w-full p-2 border rounded text-center font-bold outline-none focus:border-blue-500 text-blue-900"
                                            value={localGrades[st.id]?.[p.id] || ''} onChange={e => handleGradeChange(st.id, p.id, e.target.value)} placeholder="A/B/C" />
                                    </td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
            );
        }

        if (activeInputTab === 'ekskul') {
            return (
                <table className="w-full text-left border-collapse whitespace-nowrap">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-orange-700 text-white text-sm">
                            <th className="p-3 border-b border-r border-orange-600 text-center w-12 sticky left-0 z-30 bg-orange-800">No</th>
                            <th className="p-3 border-b border-r border-orange-600 text-center w-20 bg-orange-800">NIS</th>
                            <th className="p-3 border-b border-r border-orange-600 sticky left-12 z-30 bg-orange-800">Nama Santri</th>
                            {data.extracurriculars.map(p => <th key={p.id} className="p-3 border-b border-r border-orange-600 text-center min-w-[120px]">{p.name}</th>)}
                        </tr>
                    </thead>
                    <tbody>
                        {studentsInClass.map((st, idx) => (
                            <tr key={st.id} className="border-b hover:bg-gray-50 transition-colors">
                                <td className="p-3 text-center text-gray-500 sticky left-0 bg-white border-r z-10">{idx + 1}</td>
                                <td className="p-3 text-center bg-white border-r font-semibold">{st.nis || '-'}</td>
                                <td className="p-3 font-semibold sticky left-12 bg-white border-r shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)] z-10">{st.nama}</td>
                                {data.extracurriculars.map(p => (
                                    <td key={p.id} className="p-2 border-r bg-white hover:bg-orange-50">
                                        <input type="text" dir="auto" title="Ketik nilai (A/B/C atau angka Arab/Latin)" className="w-full p-2 border rounded text-center font-bold outline-none focus:border-orange-500 text-orange-900"
                                            value={localGrades[st.id]?.[p.id] || ''} onChange={e => handleGradeChange(st.id, p.id, e.target.value)} placeholder="A/B/C" />
                                    </td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
            );
        }

        if (activeInputTab === 'catatan') {
            return (
                <table className="w-full text-left border-collapse">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-pink-700 text-white text-sm">
                            <th className="p-3 border-b border-r border-pink-600 text-center w-12">No</th>
                            <th className="p-3 border-b border-r border-pink-600 text-center w-20">NIS</th>
                            <th className="p-3 border-b border-r border-pink-600 w-48">Nama Santri</th>
                            <th className="p-3 border-b border-pink-600">Isi Catatan Wali Kelas</th>
                        </tr>
                    </thead>
                    <tbody>
                        {studentsInClass.map((st, idx) => (
                            <tr key={st.id} className="border-b hover:bg-gray-50 transition-colors">
                                <td className="p-3 text-center text-gray-500 bg-white border-r">{idx + 1}</td>
                                <td className="p-3 text-center bg-white border-r font-semibold">{st.nis || '-'}</td>
                                <td className="p-3 font-semibold bg-white border-r">{st.nama}</td>
                                <td className="p-2 bg-white">
                                    <textarea className="w-full p-2 border rounded font-medium outline-none focus:border-pink-500 text-pink-900 min-h-[60px]"
                                        value={localGrades[st.id]?.catatan_wali || ''} onChange={e => handleGradeChange(st.id, 'catatan_wali', e.target.value)} placeholder="Tulis pesan penyemangat..." />
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            );
        }
    };

    return (
        <div ref={containerRef} className="bg-white rounded-xl shadow-sm p-6 border border-gray-100 flex flex-col h-[85vh]">
            <div className="flex flex-col gap-3 mb-4">
                <div className="flex flex-wrap items-center justify-between gap-3">
                    <div className="flex flex-wrap gap-2">
                        <button onClick={() => handleZoom('in')} className="bg-slate-600 hover:bg-slate-700 text-white px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition"><ZoomIn size={16}/>+</button>
                        <button onClick={() => handleZoom('out')} className="bg-slate-600 hover:bg-slate-700 text-white px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition"><ZoomOut size={16}/>−</button>
                        <button onClick={handleResetZoom} className="bg-slate-600 hover:bg-slate-700 text-white px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition"><Maximize size={16}/>100%</button>
                        <button onClick={toggleHeaderHidden} className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition">{isHeaderHidden ? <ChevronUp size={16}/> : <ChevronDown size={16}/>}{isHeaderHidden ? 'Tampilkan Header' : 'Sembunyikan Header'}</button>
                        <button onClick={toggleFullscreen} className="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition">{isFullscreen ? <Minimize size={16}/> : <Maximize size={16}/>} {isFullscreen ? 'Keluar Fullscreen' : 'Fullscreen'}</button>
                    </div>
                    <div className="text-xs text-gray-500">Zoom: {zoomLevel}% | Header: {isHeaderHidden ? 'Tersembunyi' : 'Tampil'}</div>
                </div>
            </div>
            {!isHeaderHidden && (
            <div className="flex flex-col mb-4 shrink-0 gap-4 border-b pb-4">
                <div className="flex justify-between items-start">
                    <div className="flex gap-4 items-center bg-gray-50 p-3 rounded-xl border flex-1">
                        <div className="flex-1 flex items-center gap-2">
                            <label className="text-sm font-medium text-gray-700 whitespace-nowrap">Pilih Kelas:</label>
                            <select className="w-full p-2 border rounded-lg focus:ring-2 outline-none font-bold" value={selectedClass} onChange={e => setSelectedClass(e.target.value)}><option value="">-- Kelas --</option>{data.classes.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select>
                        </div>
                        <div className="px-4 border-l border-r"><p className="text-xs text-gray-500">Tahun Ajaran</p><p className="font-bold text-gray-800">{activeSetting.tahun}</p></div>
                        <div className="px-4"><p className="text-xs text-gray-500">Semester</p><p className="font-bold text-gray-800">{activeSetting.semester}</p></div>
                    </div>
                    <div className="flex flex-col items-end gap-2 ml-4">
                        {isSaving ? <span className="text-xs font-bold text-yellow-600 animate-pulse bg-yellow-50 px-2 py-1 rounded-md">Menyimpan...</span> : lastSaved ? <span className="text-xs font-medium text-emerald-600 bg-emerald-50 px-2 py-1 rounded-md">Tersimpan otomatis</span> : null}
                        <button onClick={handleManualSave} disabled={!selectedClass} className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 disabled:opacity-50 transition"><Save size={16}/> Simpan Manual</button>
                    </div>
                </div>
                {/* Excel Import/Export Buttons */}
                {selectedClass && (
                    <div className="flex gap-3 items-center bg-blue-50 p-3 rounded-xl border border-blue-200">
                        <button 
                            onClick={handleExportGrades}
                            disabled={!selectedClass || subjectsInClass.length === 0}
                            className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 disabled:opacity-50 transition"
                        >
                            <Download size={16}/> Unduh Template Excel
                        </button>
                        <label className="bg-amber-600 hover:bg-amber-700 text-white px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 cursor-pointer disabled:opacity-50 transition">
                            <Upload size={16}/> {isImporting ? 'Mengimpor...' : 'Impor dari Excel'}
                            <input 
                                type="file" 
                                accept=".xlsx,.xls" 
                                className="hidden" 
                                onChange={handleImportGrades}
                                disabled={!selectedClass || subjectsInClass.length === 0 || isImporting}
                            />
                        </label>
                        <p className="text-xs text-gray-600 ml-auto">💡 Unduh template, isi nilainya, lalu impor kembali untuk update massal</p>
                    </div>
                )}
            </div>
            )}

            {selectedClass ? (
                <div className="overflow-auto border rounded-xl flex-1 relative custom-scrollbar" style={{ zoom: `${zoomLevel}%`, transformOrigin: 'top left' }}>
                    {renderTableContent()}
                </div>
            ) : (
                <div className="flex-1 flex flex-col items-center justify-center text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200"><BookOpen size={48} className="mb-4 opacity-50" /><p>Silakan pilih Kelas di atas.</p></div>
            )}
        </div>
    );
};

// ==========================================
// CETAK RAPORT / IJAZAH
// ==========================================
const CetakDokumen = ({ mode = 'raport' }) => {
    const { data, allData, addLog } = useContext(AppContext);
    const [selectedClass, setSelectedClass] = useState('');
    const [selectedStudent, setSelectedStudent] = useState('');
    const [useKatrol, setUseKatrol] = useState(false);
    const [isBatchMode, setIsBatchMode] = useState(false);
    const [printMargins, setPrintMargins] = useState({ top: 0, bottom: 0, left: 0, right: 0 });
    const [printScale, setPrintScale] = useState(0.9);
    const [previewZoom, setPreviewZoom] = useState(0.7);
    
    const activeSetting = data.settings.find(s => s.isActive) || {};
    const activeStudents = getStudentsForYear(data.studentSnapshots, activeSetting, data.students);
    const classesData = allData?.classes || data.classes;
    const studentsInClass = getStudentsInClass(activeStudents, classesData, selectedClass);
    const studentData = activeStudents.find(s => s.id === selectedStudent);
    
    const [selectedLayout, setSelectedLayout] = useState(() => {
        // Filter layouts based on mode: ijazah gets 'ijazah' layouts, raport gets non-ijazah layouts
        const available = data.layouts.filter(l =>
            mode === 'ijazah'
                ? (l.name || l.id).toLowerCase().includes('ijazah')
                : !(l.name || l.id).toLowerCase().includes('ijazah')
        );
        const matched = available.find(l => l.id === mode);
        return matched ? matched.id : (available[0]?.id || data.layouts[0]?.id || '');
    });
    
    // Available layouts filtered by mode
    const availableLayouts = data.layouts.filter(l =>
        mode === 'ijazah'
            ? (l.name || l.id).toLowerCase().includes('ijazah')
            : !(l.name || l.id).toLowerCase().includes('ijazah')
    );
    
    const layoutSettings = data.layouts.find(l => l.id === selectedLayout) || {};
    const activeLayout = layoutSettings.elements || [];
    const layoutPageSize = layoutSettings.pageSize || 'A4';
    const layoutOrientation = layoutSettings.orientation || 'portrait';
    const baseWidth = pageDimensions[layoutPageSize].width;
    const baseHeight = pageDimensions[layoutPageSize].height;
    const canvasWidth = layoutOrientation === 'landscape' ? baseHeight : baseWidth;
    const canvasHeight = layoutOrientation === 'landscape' ? baseWidth : baseHeight;

    useEffect(() => {
        setPrintMargins(layoutSettings.margins || { top: 0, bottom: 0, left: 0, right: 0 });
    }, [JSON.stringify(layoutSettings.margins)]);

    useEffect(() => {
        data.fonts?.filter(f => f.url).forEach(f => {
            const linkId = `font-link-${btoa(f.url).replace(/[^a-zA-Z0-9]/g, '')}`;
            if (!document.getElementById(linkId)) {
                const link = document.createElement('link');
                link.id = linkId;
                link.rel = 'stylesheet';
                link.href = f.url;
                document.head.appendChild(link);
            }
        });
    }, [data.fonts]);

    const gradeDocId = getGradeDocId(selectedClass, classesData, activeSetting, data.grades);
    const classGradesDoc = data.grades.find(g => g.id === gradeDocId)?.data || {};
    const studentGrades = classGradesDoc[selectedStudent] || {};

    const classAverages = useMemo(() => {
        if(!gradeDocId) return {};
        const sums = {}; const counts = {};
        Object.values(classGradesDoc).forEach(sGrades => {
            Object.entries(sGrades).forEach(([k, v]) => {
                let num = null;
                if (v && typeof v === 'object') {
                    const r = computeRaportScore(v.uts, v.uas);
                    if (r !== '') num = Number(r);
                } else if (v !== undefined && v !== '' && !isNaN(v)) {
                    num = Number(v);
                }
                if(num !== null) { sums[k] = (sums[k]||0) + num; counts[k] = (counts[k]||0) + 1; }
            });
        });
        const avgs = {};
        Object.keys(sums).forEach(k => avgs[k] = (sums[k]/counts[k]).toFixed(2));
        return avgs;
    }, [classGradesDoc, gradeDocId]);

    const [isExporting, setIsExporting] = useState(false);

    const handlePrint = () => { 
        document.title = "Cetak_Dokumen"; 
        addLog(`Mencetak ${mode} untuk ${studentData?.nama || 'Siswa'}`);
        window.print(); 
    };

    const handleSavePDF = async () => {
        if(!studentData) return;
        const ts = activeSetting.tahun ? activeSetting.tahun.replace(/\//g, '-') : 'tahun';
        const ss = activeSetting.semester || '1';
        const ns = studentData.nama.replace(/\s+/g, '_');
        
        const filename = mode === 'raport' 
            ? `raport_${ns}_${ts}_${ss}.pdf` 
            : `ijazah_${ns}_${ts}.pdf`;
            
        addLog(`Menyimpan ${mode} sebagai PDF untuk ${studentData.nama}`);
        
        const oldZoom = previewZoom;
        setPreviewZoom(1.0);
        setIsExporting(true);
        
        // Wait for re-render at scale 1.0 with no gaps
        await new Promise(r => setTimeout(r, 600));
        
        const pageContainers = document.querySelectorAll('.print-container');
        if (!pageContainers || pageContainers.length === 0) {
            setPreviewZoom(oldZoom);
            setIsExporting(false);
            return;
        }
        
        // Determine PDF page size in mm
        const isF4 = layoutPageSize === 'F4';
        const isLandscape = layoutOrientation === 'landscape';
        const pdfW = isF4 ? (isLandscape ? 330.2 : 215.9) : (isLandscape ? 297 : 210);
        const pdfH = isF4 ? (isLandscape ? 215.9 : 330.2) : (isLandscape ? 210 : 297);
        
        const pdf = new jsPDF({ unit: 'mm', format: [pdfW, pdfH], orientation: layoutOrientation });
        
        for (let i = 0; i < pageContainers.length; i++) {
            const container = pageContainers[i];
            const canvas = await html2canvas(container, {
                scale: 2,
                useCORS: true,
                logging: false,
                backgroundColor: '#ffffff',
                width: container.offsetWidth,
                height: container.offsetHeight,
                scrollX: 0,
                scrollY: 0,
                x: 0,
                y: 0,
            });
            
            if (i > 0) pdf.addPage([pdfW, pdfH], layoutOrientation);
            
            const imgData = canvas.toDataURL('image/jpeg', 0.98);
            pdf.addImage(imgData, 'JPEG', 0, 0, pdfW, pdfH);
        }
        
        pdf.save(filename);
        setPreviewZoom(oldZoom);
        setIsExporting(false);
    };

    const handleWA = () => {
        if(!studentData) return;
        const className = getClassNameFromValue(classesData, selectedClass);
        const tahun = activeSetting.tahun || '-';
        const semester = activeSetting.semester || '-';

        // Hitung ringkasan nilai per mapel (ambil max 10 mapel)
        const relevantSubjects = data.subjects.filter(s =>
            isSubjectVisibleInClass(s, selectedClass, classesData)
        ).slice(0, 10);

        const gradeLines = relevantSubjects.map(s => {
            const gradeObj = studentGrades[s.id];
            let val = '-';
            if (gradeObj && typeof gradeObj === 'object') {
                const raport = computeRaportScore(gradeObj.uts, gradeObj.uas);
                val = raport !== '' ? raport : '-';
            } else if (gradeObj !== undefined && gradeObj !== '') {
                val = gradeObj;
            }
            return `\u2022 ${s.nameId}: *${val}*`;
        }).join('\n');

        // Hitung rata-rata
        let totalVal = 0; let countVal = 0;
        relevantSubjects.forEach(s => {
            const gradeObj = studentGrades[s.id];
            let num = null;
            if (gradeObj && typeof gradeObj === 'object') {
                const r = computeRaportScore(gradeObj.uts, gradeObj.uas);
                num = r !== '' ? Number(r) : null;
            } else if (gradeObj !== undefined && gradeObj !== '' && !isNaN(gradeObj)) {
                num = Number(gradeObj);
            }
            if (num !== null) { totalVal += num; countVal++; }
        });
        const avgVal = countVal > 0 ? (totalVal / countVal).toFixed(2) : '-';

        const text = `\uD83C\uDF93 *Laporan Nilai ${mode === 'raport' ? 'Raport' : 'Ijazah'}*\nPonpes Imam Syafi'i Brebes\n\nAssalamu'alaikum Wr. Wb.\n\nDengan hormat, berikut adalah informasi nilai ananda:\n\nNama: *${studentData.nama}*\nKelas: *${className}*\nTA: *${tahun} | Semester ${semester}*\n\n\uD83D\uDCDA *Nilai Mata Pelajaran:*\n${gradeLines}\n\n\uD83D\uDCCA Rata-Rata: *${avgVal}*\n\nSemoga nilai ini menjadi motivasi untuk terus belajar. Silakan hubungi sekolah untuk pengambilan berkas fisik.\n\nWassalamu'alaikum Wr. Wb. \uD83E\uDD32`;
        addLog(`Membagikan Info ${mode} via WA untuk ${studentData.nama}`);
        window.open(`https://wa.me/?text=${encodeURIComponent(text)}`, '_blank');
    };

    const handleBatchSavePDF = () => {
        if(!selectedClass || studentsInClass.length === 0) return;
        setIsBatchMode(true);
        setTimeout(() => {
            const originalTitle = document.title;
            const ts = activeSetting.tahun ? activeSetting.tahun.replace(/\//g, '-') : 'tahun';
            const ss = activeSetting.semester || '1';
            const ks = getClassNameFromValue(classesData, selectedClass).replace(/\s+/g, '_');
            document.title = mode === 'raport' ? `raport_masal_${ts}_${ss}_${ks}` : `ijazah_masal_${ts}_${ks}`;
            addLog(`Mencetak massal ${mode} untuk kelas ${ks}`);
            window.print();
            setTimeout(() => { 
                document.title = originalTitle; 
                setIsBatchMode(false);
            }, 2000);
        }, 1000);
    };

    const getStyles = (el) => ({ position: 'absolute', left: `${el.x}px`, top: `${el.y}px`, fontSize: `${el.fontSize}px`, fontFamily: el.fontFamily || 'Arial, sans-serif', fontWeight: el.fontWeight, color: 'black', textAlign: el.textAlign || 'left' });

    const studentsToRender = isBatchMode ? studentsInClass : (studentData ? [studentData] : []);

    const renderElementForStudent = (el, stdData) => {
        const sGrades = classGradesDoc[stdData.id] || {};
        const className = getClassNameFromValue(classesData, selectedClass);
        const classDataObj = classesData.find(c => c.id === selectedClass);
        
        const replaceVariables = (str) => {
            if (typeof str !== 'string') return str;
            let replaced = str.replace(/\{\{nama_santri\}\}/gi, stdData.nama || '')
                             .replace(/\{\{nama_santri_ar\}\}/gi, stdData.nama_arab || '')
                             .replace(/\{\{nis\}\}/gi, stdData.nis || '')
                             .replace(/\{\{nisn\}\}/gi, stdData.nisn || '')
                             .replace(/\{\{kelas\}\}/gi, className || '')
                             .replace(/\{\{kelas_ar\}\}/gi, classDataObj?.name_arab || '')
                             .replace(/\{\{tahun_ajaran\}\}/gi, activeSetting.tahun || '')
                             .replace(/\{\{tahun_ajaran_ar\}\}/gi, activeSetting.tahun_arab || '')
                             .replace(/\{\{semester\}\}/gi, activeSetting.semester || '')
                             .replace(/\{\{semester_ar\}\}/gi, activeSetting.semester_arab || '');
            
            return replaced.replace(/\{\{([^}]+)\}\}/g, (match, key) => {
                 if (stdData[key] !== undefined) return stdData[key];
                 if (stdData.fields && stdData.fields[key] !== undefined) return stdData.fields[key];
                 return match;
            });
        };

        let content = replaceVariables(el.content);
        
        const baseStyle = {
            position: 'absolute',
            left: `${el.x}px`,
            top: `${el.y}px`,
            fontSize: `${el.fontSize}px`,
            fontFamily: el.fontFamily || 'Arial, sans-serif',
            fontWeight: el.fontWeight,
            color: 'black',
            zIndex: el.zIndex ?? 1,
            opacity: el.opacity ?? 1,
            textAlign: el.textAlign || 'left',
        };

        if (el.type === 'table_grades') {
            return (
                <div style={{
                    ...baseStyle,
                    width: el.width ? `${el.width}px` : 'auto',
                    height: el.height ? `${el.height}px` : 'auto',
                    padding: 0,
                    background: el.isTransparent ? 'transparent' : 'white',
                    overflow: 'hidden',
                }}>
                    {renderDynamicTable(el, data, sGrades, classAverages, useKatrol, mode, classesData)}
                </div>
            );
        }
        if (el.type === 'image') {
            return (
                <img
                    src={el.content}
                    style={{
                        ...baseStyle,
                        width: `${el.width}px`,
                        height: `${el.height}px`,
                        objectFit: el.objectFit || 'contain',
                        objectPosition: `${el.objectPositionX ?? 50}% ${el.objectPositionY ?? 50}%`,
                    }}
                    alt="elemen"
                />
            );
        }
        if (el.type === 'group') {
            return (
                <div style={{ ...baseStyle, width: `${el.width}px`, height: `${el.height}px` }}>
                    {(el.children || []).map(child => {
                        const childStyle = {
                            position: 'absolute',
                            left: `${child.x}px`, top: `${child.y}px`,
                            fontSize: `${child.fontSize}px`,
                            fontFamily: child.fontFamily || 'Arial, sans-serif',
                            fontWeight: child.fontWeight,
                            color: 'black',
                            zIndex: child.zIndex ?? 1,
                            opacity: child.opacity ?? 1,
                            width: child.width ? `${child.width}px` : 'auto',
                            height: (child.type === 'image' || child.type === 'table_grades' || child.type === 'shape') ? (child.height ? `${child.height}px` : 'auto') : child.type === 'line' ? `${child.lineThickness || 2}px` : 'auto',
                            padding: (child.type === 'image' || child.type === 'table_grades' || child.type === 'shape' || child.type === 'line') ? 0 : '2px',
                        };
                        if (child.type === 'table_grades') return <div key={child.id} style={{...childStyle, background: child.isTransparent ? 'transparent' : 'white', overflow:'hidden'}}>{renderDynamicTable(child, data, sGrades, classAverages, useKatrol, mode, classesData)}</div>;
                        if (child.type === 'image') return <img key={child.id} src={child.content} style={{...childStyle, objectFit: child.objectFit||'contain', objectPosition:`${child.objectPositionX??50}% ${child.objectPositionY??50}%`}} alt="c" />;
                        if (child.type === 'line') return <div key={child.id} style={{...childStyle, backgroundColor: child.lineColor || '#000000'}} />;
                        if (child.type === 'shape') return <div key={child.id} style={{...childStyle, backgroundColor: child.shapeFill || '#000000', borderRadius: `${child.shapeRadius || 0}px`, border: child.shapeBorder ? `${child.shapeBorder}px solid ${child.shapeBorderColor || '#000000'}` : 'none'}} />;
                        return <div key={child.id} style={{...childStyle, whiteSpace:'pre-wrap', direction: child.isRtl ? 'rtl' : 'ltr'}}>{replaceVariables(child.content)}</div>;
                    })}
                </div>
            );
        }
        if (el.type === 'line') {
            return <div style={{ ...baseStyle, width: `${el.width}px`, height: `${el.lineThickness || 2}px`, backgroundColor: el.lineColor || '#000000' }} />;
        }
        if (el.type === 'shape') {
            return <div style={{ ...baseStyle, width: `${el.width}px`, height: `${el.height}px`, backgroundColor: el.shapeFill || '#000000', borderRadius: `${el.shapeRadius || 0}px`, border: el.shapeBorder ? `${el.shapeBorder}px solid ${el.shapeBorderColor || '#000000'}` : 'none' }} />;
        }
        // text / default
        return (
            <div style={{ ...baseStyle, whiteSpace: 'pre-wrap', width: el.width ? `${el.width}px` : 'auto', direction: el.isRtl ? 'rtl' : 'ltr' }}>
                {content}
            </div>
        );
    };

    const cssPageSize = layoutPageSize === 'F4' ? `215.9mm 330.2mm ${layoutOrientation}` : `A4 ${layoutOrientation}`;
    const maxPage = activeLayout.length > 0 ? Math.max(...activeLayout.map(el => el.pageIndex || 0)) : 0;
    const pages = Array.from({length: maxPage + 1}, (_, i) => activeLayout.filter(el => (el.pageIndex || 0) === i));

    return (
        <div className="flex flex-col md:flex-row gap-6 flex-1 h-full min-h-0">
            <div className="w-full md:w-80 bg-white p-6 rounded-xl shadow-sm border border-gray-100 print:hidden shrink-0 h-full overflow-y-auto">
                <h3 className="text-xl font-bold mb-4 capitalize">Cetak {mode}</h3>
                {!activeSetting.tahun && <div className="bg-yellow-50 text-yellow-800 p-3 rounded-lg text-sm mb-4 border border-yellow-200">Pastikan Admin mengaktifkan Tahun Ajaran di Master Data terlebih dahulu.</div>}
                <div className="space-y-4">
                    <select className="w-full p-2 border rounded-lg" value={selectedClass} onChange={e => {setSelectedClass(e.target.value); setSelectedStudent('');}}><option value="">-- Kelas --</option>{data.classes.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select>
                    <select className="w-full p-2 border rounded-lg" value={selectedStudent} onChange={e => setSelectedStudent(e.target.value)} disabled={!selectedClass}><option value="">-- Santri --</option>{studentsInClass.map(s => <option key={s.id} value={s.id}>{s.nama}</option>)}</select>
                    
                    {/* Pilihan Layout */}
                    <div className="pt-2 border-t">
                        <p className="text-xs font-bold text-gray-700 mb-1">Layout yang Digunakan</p>
                        <select className="w-full p-2 border rounded-lg bg-white text-sm font-semibold text-emerald-800" value={selectedLayout} onChange={e => setSelectedLayout(e.target.value)}>
                            {availableLayouts.map(l => <option key={l.id} value={l.id}>{l.name || l.id}</option>)}
                        </select>
                        {activeLayout.length === 0 && <p className="text-xs text-red-500 mt-1">⚠ Layout yang dipilih belum memiliki elemen. Silakan desain di Layout Builder terlebih dahulu.</p>}
                    </div>
                    <div className="pt-4 border-t"><label className="flex items-center gap-2 bg-yellow-50 p-3 rounded-lg border border-yellow-200 cursor-pointer"><input type="checkbox" className="w-5 h-5 text-yellow-600" checked={useKatrol} onChange={e => setUseKatrol(e.target.checked)} /><div><p className="font-bold text-yellow-800 text-sm">Gunakan Nilai Katrol</p><p className="text-xs text-yellow-700">Nilai merah otomatis menjadi KKM</p></div></label></div>
                    <div className="pt-4 border-t">
                        <p className="text-sm font-bold text-gray-700 mb-2">Penyesuaian Margin Printer (mm)</p>
                        <div className="grid grid-cols-4 gap-2">
                            <div className="flex flex-col"><label className="text-[10px] text-gray-500 mb-1">Atas</label><input type="number" value={printMargins.top} onChange={e=>setPrintMargins({...printMargins, top: Number(e.target.value)})} className="w-full p-2 border rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-gray-500 mb-1">Bawah</label><input type="number" value={printMargins.bottom} onChange={e=>setPrintMargins({...printMargins, bottom: Number(e.target.value)})} className="w-full p-2 border rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-gray-500 mb-1">Kiri</label><input type="number" value={printMargins.left} onChange={e=>setPrintMargins({...printMargins, left: Number(e.target.value)})} className="w-full p-2 border rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-gray-500 mb-1">Kanan</label><input type="number" value={printMargins.right} onChange={e=>setPrintMargins({...printMargins, right: Number(e.target.value)})} className="w-full p-2 border rounded text-xs" /></div>
                        </div>
                        <p className="text-[10px] text-gray-400 mt-1 leading-tight">Margin awal diambil dari master Layout. Setting ini akan menggeser cetakan secara fisik di kertas.</p>
                    </div>
                    <div className="pt-4 border-t">
                        <p className="text-sm font-bold text-gray-700 mb-1">Skala Cetak (Print Scale)</p>
                        <div className="flex items-center gap-3">
                            <input type="range" min="0.5" max="1.0" step="0.01" value={printScale} onChange={e => setPrintScale(Number(e.target.value))} className="flex-1 h-2 accent-blue-600" />
                            <div className="flex items-center gap-1 shrink-0">
                                <input type="number" min="50" max="100" value={Math.round(printScale * 100)} onChange={e => setPrintScale(Math.min(1, Math.max(0.5, Number(e.target.value) / 100)))} className="w-14 p-1.5 border rounded text-xs text-center font-bold" />
                                <span className="text-xs text-gray-500">%</span>
                            </div>
                        </div>
                        <p className="text-[10px] text-gray-400 mt-1 leading-tight">Default 90% karena DPI printer fisik berbeda dengan layar. Sesuaikan sampai hasil cetak pas di kertas.</p>
                    </div>
                    <div className="pt-4 flex flex-col gap-3">
                        <button onClick={handlePrint} disabled={!selectedStudent} className="w-full bg-blue-600 hover:bg-blue-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition"><Printer size={18}/> Print Langsung</button>
                        <button onClick={handleSavePDF} disabled={!selectedStudent} className="w-full bg-red-600 hover:bg-red-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition"><Download size={18}/> Simpan sbg PDF</button>
                        <button onClick={handleWA} disabled={!selectedStudent} className="w-full bg-emerald-600 hover:bg-emerald-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition"><Share2 size={18}/> Kirim Info via WA</button>
                    </div>
                    <div className="pt-2 border-t mt-4">
                        <button onClick={handleBatchSavePDF} disabled={!selectedClass || studentsInClass.length === 0} className="w-full bg-orange-600 hover:bg-orange-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition"><Printer size={18}/> Cetak Semua (1 Kelas)</button>
                        <p className="text-xs text-gray-500 text-center mt-2">Cetak raport seluruh siswa di kelas yang dipilih dalam 1 file PDF.</p>
                    </div>
                </div>
            </div>
            <div className="flex-1 bg-gray-200 rounded-xl overflow-auto border border-gray-300 print:bg-white print:p-0 print:border-none print:overflow-visible relative print:static flex flex-col">
                {/* Zoom Controls Bar */}
                <div className="print:hidden sticky top-0 z-20 bg-white/90 backdrop-blur border-b border-gray-200 px-4 py-2 flex items-center gap-3 shrink-0">
                    <span className="text-xs font-bold text-gray-600">Zoom Preview:</span>
                    <button onClick={() => setPreviewZoom(z => Math.max(0.3, +(z - 0.1).toFixed(1)))} className="bg-gray-100 hover:bg-gray-200 rounded px-2 py-1 text-sm font-bold">−</button>
                    <span className="text-xs font-mono w-10 text-center">{Math.round(previewZoom * 100)}%</span>
                    <button onClick={() => setPreviewZoom(z => Math.min(1.5, +(z + 0.1).toFixed(1)))} className="bg-gray-100 hover:bg-gray-200 rounded px-2 py-1 text-sm font-bold">+</button>
                    <button onClick={() => setPreviewZoom(0.7)} className="text-xs text-blue-600 hover:underline ml-1">Reset</button>
                    <button onClick={() => setPreviewZoom(1.0)} className="text-xs text-gray-500 hover:underline">100%</button>
                </div>
                <div className="flex-1 overflow-auto p-8 flex flex-col items-center gap-8 print:p-0 print:gap-0 print:block">
                {studentsToRender.length > 0 ? (
                    <div className={`print-wrapper flex flex-col print:block ${isExporting ? 'gap-0 items-start' : 'items-center gap-8 print:gap-0'}`} style={{ transformOrigin: 'top left', transform: isExporting ? 'none' : `scale(${previewZoom})`, marginBottom: isExporting ? '0px' : `${(previewZoom - 1) * canvasHeight * pages.length}px` }}>
                        {studentsToRender.map((std, stdIndex) => (
                            <React.Fragment key={std.id}>
                                {pages.map((pageElements, index) => {
                                    const isLastPageOfLastStudent = stdIndex === studentsToRender.length - 1 && index === pages.length - 1;
                                    return (
                                        <div key={`${std.id}-${index}`} className={`print-container bg-white relative print:shadow-none print:!m-0 ${isExporting ? '!m-0 shadow-none' : 'shadow-xl'}`} style={{ width: `${canvasWidth}px`, height: `${canvasHeight}px`, pageBreakAfter: isLastPageOfLastStudent ? 'auto' : 'always', marginTop: (index === 0 && stdIndex === 0) || isExporting ? 0 : '32px', flexShrink: 0 }}>
                                            {pageElements.length > 0 ? pageElements.map(el => <React.Fragment key={el.id}>{renderElementForStudent(el, std)}</React.Fragment>) : (activeLayout.length === 0 && index === 0 ? <div className="absolute inset-0 flex items-center justify-center text-gray-400 print:hidden">Layout belum disetting oleh Admin.<br/>Pilih layout di panel kiri.</div> : <div className="absolute inset-0 flex items-center justify-center text-gray-400 print:hidden">Halaman {index + 1} kosong.</div>)}
                                        </div>
                                    );
                                })}
                            </React.Fragment>
                        ))}
                    </div>
                ) : <div className="flex items-center justify-center h-full text-gray-400 print:hidden">Pilih santri untuk melihat preview {mode}.</div>}
                </div>
            </div>
            <style>{`
            @media print { 
                body * { visibility: hidden; } 
                .print-wrapper, .print-wrapper * { visibility: visible; } 
                .print-wrapper { 
                    position: absolute !important; 
                    left: 0 !important; 
                    top: 0 !important; 
                    margin: 0 !important; 
                    padding: 0 !important;
                    transform: scale(${printScale}) !important;
                    transform-origin: top left !important;
                    width: ${Math.round(100 / printScale)}% !important;
                } 
                .print-container { position: relative !important; padding: 0 !important; box-shadow: none !important; border: none !important; transform: none !important; left: 0 !important; top: 0 !important; margin: 0 !important; } 
                @page { size: ${cssPageSize}; margin: 0 !important; } 
            }
            `}</style>
        </div>
    );
};

// ==========================================
// LEGER KELAS
// ==========================================
const LeggerKelas = () => {
    const { data, allData, addLog } = useContext(AppContext);
    const [selectedClass, setSelectedClass] = useState('');
    const classesData = allData?.classes || data.classes;
    
    const activeSetting = data.settings.find(s => s.isActive) || {};
    const activeStudents = useMemo(() => getStudentsForYear(data.studentSnapshots, activeSetting, data.students), [data.studentSnapshots, activeSetting, data.students]);
    const students = useMemo(() => getStudentsInClass(activeStudents, classesData, selectedClass), [activeStudents, classesData, selectedClass]);
    const subjects = useMemo(() => sortSubjectsByCategory(filterSubjectsByClass(data.subjects, selectedClass, classesData), data.subjectCategories), [data.subjects, data.subjectCategories, selectedClass, classesData]);
    
    const gradeDocId = getGradeDocId(selectedClass, classesData, activeSetting, data.grades);

    const grades = data.grades.find(g => g.id === gradeDocId)?.data || {};

    const getSubjectGradeValue = (grade) => {
        if (grade === null || grade === undefined || grade === '') return null;
        if (typeof grade === 'object') {
            if ('raport' in grade) {
                const raportValue = parseGradeValue(grade.raport);
                if (raportValue !== null) return raportValue;
            }
            const raportScore = computeRaportScore(grade.uts, grade.uas);
            return raportScore !== '' ? raportScore : null;
        }
        return parseGradeValue(grade);
    };

    const formatSubjectGradeDisplay = (grade) => {
        if (grade === null || grade === undefined || grade === '') return '-';
        if (typeof grade === 'object') {
            const uts = grade.uts ?? '';
            const uas = grade.uas ?? '';
            const raport = grade.raport ?? '';
            if (raport !== '') {
                const raportValue = parseGradeValue(raport);
                if (raportValue !== null) return raportValue;
            }
            if (uts !== '' || uas !== '') {
                const formattedUts = uts === '' ? '-' : uts;
                const formattedUas = uas === '' ? '-' : uas;
                const raportScore = computeRaportScore(uts, uas);
                return raportScore !== '' ? `${formattedUts}/${formattedUas} (${raportScore})` : `${formattedUts}/${formattedUas}`;
            }
            return '-';
        }
        return parseGradeValue(grade) ?? '-';
    };

    const leggerData = useMemo(() => {
        if (!selectedClass) return [];
        let r = students.map(st => {
            let total = 0; let count = 0;
            subjects.forEach(sub => {
                const value = getSubjectGradeValue(grades[st.id]?.[sub.id]);
                if (value !== null && !isNaN(value)) { total += value; count++; }
            });
            const avg = count > 0 ? (total / count).toFixed(2) : 0;
            let predikat = 'D'; if (avg >= 90) predikat = 'A'; else if (avg >= 80) predikat = 'B'; else if (avg >= 70) predikat = 'C';
            return { ...st, total, avg, predikat, grades: grades[st.id] || {} };
        });
        r.sort((a, b) => b.avg - a.avg); return r;
    }, [students, subjects, grades, selectedClass]);

    const handleSendWA = (row) => {
        const className = getClassNameFromValue(classesData, selectedClass);
        const tahun = activeSetting.tahun || '-';
        const semester = activeSetting.semester || '-';
        const gradeLines = subjects.slice(0, 10).map(s => {
            const val = getSubjectGradeValue(row.grades[s.id]);
            return `\u2022 ${s.nameId}: *${val !== null ? val : '-'}*`;
        }).join('\n');
        const text = `\uD83C\uDF93 *Laporan Nilai Santri*\nPonpes Imam Syafi'i Brebes\n\nNama: *${row.nama}*\nKelas: *${className}*\nTA: *${tahun} Sem ${semester}*\n\n\uD83D\uDCDA *Ringkasan Nilai:*\n${gradeLines}\n\n\uD83D\uDCCA Rata-Rata: *${row.avg}* | Predikat: *${row.predikat}*\n\nAlhamdulillah, terima kasih atas kepercayaan Anda. \uD83E\uDD32`;
        addLog(`Kirim info nilai ${row.nama} via WA`);
        window.open(`https://wa.me/?text=${encodeURIComponent(text)}`, '_blank');
    };

    const exportExcel = () => {
        const className = getClassNameFromValue(classesData, selectedClass);
        addLog(`Mengekspor Legger Kelas ${className}`);
        
        const categoryHeaders = ['No.', 'NIS', 'Nama Santri'];
        const headers = ['', '', ''];
        
        let currentColIndex = 3;
        const merges = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: subjects.length + 5 } },
            { s: { r: 1, c: 0 }, e: { r: 1, c: subjects.length + 5 } },
            { s: { r: 3, c: 0 }, e: { r: 4, c: 0 } }, // Peringkat vertikal
            { s: { r: 3, c: 1 }, e: { r: 4, c: 1 } }, // NIS vertikal
            { s: { r: 3, c: 2 }, e: { r: 4, c: 2 } }  // Nama Santri vertikal
        ];

        Object.entries(groupBy(subjects, 'kategori')).forEach(([cat, subs]) => {
            categoryHeaders.push(cat || 'Umum');
            for (let i = 1; i < subs.length; i++) categoryHeaders.push('');
            if (subs.length > 1) {
                merges.push({ s: { r: 3, c: currentColIndex }, e: { r: 3, c: currentColIndex + subs.length - 1 } });
            }
            subs.forEach(s => headers.push(s.nameId));
            currentColIndex += subs.length;
        });

        categoryHeaders.push('Total', 'Rata-rata', 'Predikat');
        headers.push('', '', '');
        
        merges.push(
            { s: { r: 3, c: currentColIndex }, e: { r: 4, c: currentColIndex } },
            { s: { r: 3, c: currentColIndex + 1 }, e: { r: 4, c: currentColIndex + 1 } },
            { s: { r: 3, c: currentColIndex + 2 }, e: { r: 4, c: currentColIndex + 2 } }
        );
        
        const aoa = [
            [`DATA NILAI LEGGER KELAS ${className.toUpperCase()}`],
            [`Tahun Ajaran: ${activeSetting?.tahun || '-'} | Semester: ${activeSetting?.semester || '-'}`],
            [],
            categoryHeaders,
            headers
        ];
        
        leggerData.forEach((row, idx) => {
            const rowData = [idx + 1, row.nis || '', row.nama || ''];
            subjects.forEach(s => {
                const raw = row.grades[s.id];
                const value = formatSubjectGradeDisplay(raw);
                rowData.push(value);
            });
            rowData.push(row.total, row.avg, row.predikat);
            aoa.push(rowData);
        });

        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws['!merges'] = merges;

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, `Legger ${className}`);
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Legger_Kelas_${className.replace(/\s+/g, '_')}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
    };

    return (
        <div className="bg-white rounded-xl shadow-sm p-6 border border-gray-100 flex flex-col h-[85vh]">
            <div className="flex justify-between items-center mb-6 shrink-0">
                <h2 className="text-2xl font-bold text-gray-800 flex items-center gap-2"><BookOpen /> Legger Kelas</h2>
                <div className="flex gap-4">
                    <select className="p-2 border rounded-lg min-w-[150px]" value={selectedClass} onChange={e => setSelectedClass(e.target.value)}><option value="">-- Pilih Kelas --</option>{data.classes.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select>
                    <button onClick={exportExcel} disabled={!selectedClass} className="bg-emerald-600 text-white px-4 py-2 rounded-lg font-medium flex items-center gap-2 disabled:opacity-50"><Download size={18}/> Ekspor Excel</button>
                </div>
            </div>
            {selectedClass ? (
                <div className="overflow-auto border rounded-xl flex-1 relative">
                    <table className="w-full text-left border-collapse whitespace-nowrap">
                        <thead className="sticky top-0 z-20">
                            <tr className="bg-gray-800 text-white text-sm">
                                <th rowSpan={3} className="p-3 border-b border-gray-700 text-center w-12 sticky left-0 z-30 bg-gray-900">Rank</th>
                                <th rowSpan={3} className="p-3 border-b border-gray-700 sticky left-12 z-30 bg-gray-900">Nama Santri</th>
                                {subjects.length > 0 && <th colSpan={subjects.length} className="p-3 border-b border-gray-700 text-center bg-gray-900">Mata Pelajaran</th>}
                                <th rowSpan={3} className="p-3 border-b border-gray-700 text-center bg-gray-700">Total</th><th rowSpan={3} className="p-3 border-b border-gray-700 text-center bg-gray-700">Rata-rata</th><th rowSpan={3} className="p-3 border-b border-gray-700 text-center bg-gray-700">Predikat</th><th rowSpan={3} className="p-3 border-b border-gray-700 text-center bg-emerald-700 print:hidden">WA</th>
                            </tr>
                            <tr className="bg-gray-700 text-white text-sm">
                                {Object.entries(groupBy(subjects, 'kategori')).map(([cat, subs]) => (
                                    <th key={cat || 'umum'} colSpan={subs.length} className="p-3 border-b border-gray-700 text-center bg-gray-800">{cat || 'Umum'}</th>
                                ))}
                            </tr>
                            <tr className="bg-gray-800 text-white text-sm">
                                {subjects.map(s => <th key={s.id} className="p-3 border-b border-gray-700 text-center" title={s.nameId}><div className="w-20 truncate">{s.nameId}</div></th>)}
                            </tr>
                        </thead>
                        <tbody>
                            {leggerData.map((row, idx) => (
                                <tr key={row.id} className="border-b hover:bg-gray-50 text-sm">
                                    <td className="p-3 text-center font-bold sticky left-0 bg-white border-r z-10">{idx + 1}</td>
                                    <td className="p-3 font-semibold sticky left-12 bg-white border-r shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)] z-10">{row.nama}</td>
                                    {subjects.map(s => {
                                        const raw = row.grades[s.id];
                                        const display = formatSubjectGradeDisplay(raw);
                                        const val = getSubjectGradeValue(raw);
                                        const isRed = val !== null && Number(val) > 0 && Number(val) < Number(s.kkm);
                                        return <td key={s.id} className={`p-3 text-center border-r ${isRed ? 'text-red-600 font-bold bg-red-50' : ''}`}>{display}</td>;
                                    })}
                                    <td className="p-3 text-center font-bold bg-emerald-50 border-r">{row.total}</td><td className="p-3 text-center font-bold bg-emerald-100 border-r">{row.avg}</td><td className="p-3 text-center font-bold bg-emerald-50">{row.predikat}</td>
                                    <td className="p-3 text-center print:hidden">
                                        <button onClick={() => handleSendWA(row)} title="Kirim Nilai via WhatsApp" className="bg-emerald-500 hover:bg-emerald-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1 mx-auto transition">
                                            <Share2 size={13}/> WA
                                        </button>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            ) : <div className="flex-1 flex items-center justify-center text-gray-400 bg-gray-50 rounded-xl border border-dashed">Pilih kelas untuk melihat legger.</div>}
        </div>
    );
};

// ==========================================
// MAIN DASHBOARD
// ==========================================
const TahunAjaranSelection = ({ allData, saveToDb, onBypass, currentUser }) => {
    const settingsList = allData?.settings || [];
    const hasSettings = settingsList.length > 0;
    
    return (
        <div className="min-h-screen bg-emerald-900 flex flex-col items-center justify-center p-6 relative overflow-hidden">
            <div className="absolute top-0 left-0 w-full h-full overflow-hidden pointer-events-none">
                <div className="absolute -top-24 -left-24 w-96 h-96 bg-emerald-800 rounded-full mix-blend-multiply blur-3xl opacity-50"></div>
                <div className="absolute top-1/2 right-1/4 w-72 h-72 bg-emerald-700 rounded-full mix-blend-multiply blur-3xl opacity-50"></div>
                <div className="absolute -bottom-32 -left-1/4 w-80 h-80 bg-emerald-600 rounded-full mix-blend-multiply blur-3xl opacity-50"></div>
            </div>

            <div className="z-10 bg-white/10 backdrop-blur-lg border border-white/20 p-8 rounded-3xl shadow-2xl max-w-3xl w-full text-center">
                <img src="https://i.ibb.co.com/DfZSFRsP/Chat-GPT-Image-3-Mei-2026-04-08-56.png" alt="Logo" className="w-24 h-24 mx-auto mb-6 drop-shadow-xl" />
                <h1 className="text-4xl font-bold text-white mb-2">Rapijaz-Maisya</h1>
                <p className="text-emerald-100 mb-8 text-lg">Silakan Pilih Tahun Ajaran Aktif untuk Melanjutkan</p>

                {hasSettings ? (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 max-h-[50vh] overflow-y-auto p-2 custom-scrollbar">
                        {settingsList.map(s => (
                            <button
                                key={s.id}
                                onClick={() => saveToDb('settings', s.id, { ...s, isActive: true }, true, `Mengaktifkan semester ${s.tahun} ${s.semester}`)}
                                className="bg-white/90 hover:bg-white text-emerald-900 rounded-2xl p-6 text-left shadow-lg hover:shadow-xl transition-all transform hover:-translate-y-1 group border-2 border-transparent hover:border-emerald-400 flex items-center justify-between"
                            >
                                <div>
                                    <h3 className="text-xl font-bold">{s.tahun}</h3>
                                    <p className="text-sm font-medium text-emerald-600">Semester {s.semester}</p>
                                </div>
                                <div className="bg-emerald-100 text-emerald-700 p-3 rounded-full opacity-0 group-hover:opacity-100 transition-opacity">
                                    <CheckCircle size={24} />
                                </div>
                            </button>
                        ))}
                    </div>
                ) : (
                    <div className="bg-white/20 p-8 rounded-2xl border border-white/30 text-white">
                        <AlertCircle size={48} className="mx-auto mb-4 text-emerald-200" />
                        <h3 className="text-xl font-bold mb-2">Belum Ada Tahun Ajaran</h3>
                        <p className="text-emerald-100">Silakan hubungi Administrator untuk mengatur Tahun Ajaran pertama kali.</p>
                    </div>
                )}

                {currentUser?.role === 'admin' && (
                    <button
                        onClick={onBypass}
                        className="mt-8 text-emerald-200 hover:text-white text-sm font-medium flex items-center gap-2 mx-auto transition-colors"
                    >
                        <Settings size={16} /> Masuk ke Master Data (Bypass Admin)
                    </button>
                )}
            </div>
        </div>
    );
};

const Dashboard = () => {
  const { currentUser, setCurrentUser, activeSetting, allData, saveToDb } = useContext(AppContext);
  const isTahunSet = !!activeSetting?.tahun;
  const [bypassSplash, setBypassSplash] = useState(false);
  const [activeMenu, setActiveMenu] = useState('dashboard');
  const [expandedMenu, setExpandedMenu] = useState('');
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isSidebarCompact, setIsSidebarCompact] = useState(false);
  const [isLogoutModalOpen, setIsLogoutModalOpen] = useState(false);

  // ==========================================
  // GLOBAL EVENT LISTENER - KONVERSI ANGKA ARAB
  // ==========================================
  useEffect(() => {
    const handleInput = (e) => {
      const target = e.target;
      // Hanya konversi untuk input[type="text"] dan textarea
      if ((target.tagName === 'INPUT' && target.type === 'text') || target.tagName === 'TEXTAREA') {
        const convertedValue = convertArabicToLatin(target.value);
        if (convertedValue !== target.value) {
          // Preservasi cursor position
          const cursorPos = target.selectionStart;
          target.value = convertedValue;
          target.selectionStart = target.selectionEnd = cursorPos;
          // Trigger onChange event untuk React
          const event = new Event('input', { bubbles: true });
          target.dispatchEvent(event);
          const changeEvent = new Event('change', { bubbles: true });
          target.dispatchEvent(changeEvent);
        }
      }
    };

    // Event listeners untuk input dan paste
    document.addEventListener('input', handleInput);
    document.addEventListener('paste', (e) => {
      setTimeout(() => handleInput({ target: e.target }), 0);
    });

    return () => {
      document.removeEventListener('input', handleInput);
      document.removeEventListener('paste', handleInput);
    };
  }, []);



  const masterDataSubItems = [
    { id: 'settings', label: 'Tahun Ajaran' }, 
    { id: 'classes', label: 'Daftar Kelas' },
    { id: 'teachers', label: 'Guru Pengampu' },
    { id: 'subjectCategories', label: 'Kategori Pelajaran' },
    { id: 'masterSubjects', label: 'Daftar Pelajaran Utama' },
    { id: 'subjects', label: 'Plotting Pelajaran' },
    { id: 'presences', label: 'Presensi' }, 
    { id: 'characterTraits', label: 'Sikap & Kesantrian' }, 
    { id: 'extracurriculars', label: 'Ekstrakurikuler' }, 
    { id: 'studentFields', label: 'Field Santri' }, 
    { id: 'students', label: 'Data Santri' },
    { id: 'fonts', label: 'Font Kustom' }, 
    { id: 'users', label: 'Pengguna Sistem' },
    { id: 'backup_restore', label: 'Backup & Restore' }
  ];

  const inputNilaiSubItems = [
    { id: 'pelajaran', label: 'Nilai Pelajaran' },
    { id: 'presensi', label: 'Presensi' },
    { id: 'sikap', label: 'Sikap & Kesantrian' },
    { id: 'ekskul', label: 'Ekstrakurikuler' },
    { id: 'catatan', label: 'Catatan Wali Kelas' }
  ];

  const menuItems = [
    { id: 'dashboard', label: 'Dashboard', icon: Home, roles: ['admin', 'guru', 'user'] },
    { id: 'master_data', label: 'Master Data', icon: Users, roles: ['admin'], subItems: masterDataSubItems },
    { id: 'layout_builder', label: 'Desain Layout', icon: LayoutTemplate, roles: ['admin'] },
    { id: 'input_nilai', label: 'Input Nilai', icon: CheckSquare, roles: ['admin', 'guru', 'user'], subItems: inputNilaiSubItems },
    { id: 'legger', label: 'Legger Kelas', icon: BookOpen, roles: ['admin', 'guru', 'user'] },
    { id: 'cetak_raport', label: 'Cetak Raport', icon: Printer, roles: ['admin', 'guru', 'user'] },
    { id: 'cetak_ijazah', label: 'Cetak Ijazah', icon: Printer, roles: ['admin'] },
  ];

  const filteredMenu = menuItems.filter(m => m.roles.includes(currentUser?.role));

  const renderContent = () => {
    if (masterDataSubItems.some(sub => sub.id === activeMenu)) {
        return <MasterData activeTab={activeMenu} />;
    }
    if (inputNilaiSubItems.some(sub => sub.id === activeMenu)) {
        return <InputNilai activeInputTab={activeMenu} />;
    }

    switch (activeMenu) {
      case 'dashboard': return <HomeDashboard />;
      case 'layout_builder': return <LayoutBuilder />;
      case 'cetak_raport': return <CetakDokumen key="raport" mode="raport" />;
      case 'cetak_ijazah': return <CetakDokumen key="ijazah" mode="ijazah" />;
      case 'legger': return <LeggerKelas />;
      default: return <div className="p-8 text-center text-gray-500">Menu tidak ditemukan</div>;
    }
  };

  const getMenuLabel = () => {
      for (let m of menuItems) {
          if (m.id === activeMenu) return m.label;
          if (m.subItems) {
              let sub = m.subItems.find(s => s.id === activeMenu);
              if (sub) return `${m.label} / ${sub.label}`;
          }
      }
      return 'Menu';
  };

  return (
    <div className="min-h-screen bg-gray-50 flex">
      {isSidebarOpen && <div className="fixed inset-0 bg-black/50 z-40 md:hidden" onClick={() => setIsSidebarOpen(false)} />}
      <div className={`fixed inset-y-0 left-0 ${isSidebarCompact ? 'w-16' : 'w-64'} bg-emerald-800 text-emerald-50 transform ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full'} md:relative md:translate-x-0 transition-transform duration-200 ease-in-out z-50 flex flex-col`}>
        <div className={`p-3 ${isSidebarCompact ? 'px-0' : 'p-6'} flex items-center justify-between border-b border-emerald-700/50`}>
          <div className={`flex items-center ${isSidebarCompact ? 'justify-center w-full gap-0' : 'gap-3'}`}>
            <img src="https://i.ibb.co.com/DfZSFRsP/Chat-GPT-Image-3-Mei-2026-04-08-56.png" alt="Logo" className="w-12 h-12 object-contain drop-shadow-sm shrink-0" />
            <div className={`${isSidebarCompact ? 'hidden' : 'font-bold text-xl leading-tight'}`}>
              Rapijaz-Maisya<br/>
              <span className="text-emerald-300 text-[10px] font-normal leading-tight block mt-1">
                Aplikasi Raport dan Ijazah<br/>Ponpes Imam Syafi'i Brebes
              </span>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button className="hidden md:inline-flex text-emerald-200 hover:text-white" onClick={() => setIsSidebarCompact(prev => !prev)} title="Ciutkan Sidebar">
              <LayoutTemplate size={20} />
            </button>
            <button className="md:hidden text-emerald-200" onClick={() => setIsSidebarOpen(false)}><X size={24}/></button>
          </div>
        </div>
        
        <div className="p-4 flex-1 overflow-y-auto custom-scrollbar">
            <div className="space-y-1">
                {filteredMenu.map(menu => { 
                    const Icon = menu.icon;
                    const isExpanded = expandedMenu === menu.id;
                    const isActive = activeMenu === menu.id || (menu.subItems && menu.subItems.some(s => s.id === activeMenu));
                    const isDisabled = !isTahunSet && menu.id !== 'master_data';

                    return (
                        <div key={menu.id} className="mb-1">
                            <button 
                                onClick={() => { 
                                    if(isDisabled) return;
                                    if(menu.subItems) {
                                        setExpandedMenu(isExpanded ? '' : menu.id);
                                        if (!isExpanded && menu.subItems.length > 0) {
                                            setActiveMenu(menu.subItems[0].id);
                                        }
                                    } else {
                                        setActiveMenu(menu.id); setIsSidebarOpen(false); setExpandedMenu('');
                                    }
                                }} 
                                title={isDisabled ? 'Atur Tahun Ajaran Dulu' : menu.label}
                                disabled={isDisabled}
                                className={`w-full flex items-center ${isSidebarCompact ? 'justify-center px-2 py-3' : 'justify-between px-4 py-3'} rounded-xl transition ${isActive && !menu.subItems ? 'bg-emerald-600 text-white shadow-lg' : 'text-emerald-100 hover:bg-emerald-700/50'} ${isDisabled ? 'opacity-50 cursor-not-allowed' : ''}`}
                            >
                                <div className={`flex items-center ${isSidebarCompact ? 'justify-center gap-0' : 'gap-3'}`}>
                                    <Icon size={20} />
                                    <span className={`${isSidebarCompact ? 'hidden' : 'font-medium'}`}>{menu.label}</span>
                                </div>
                                {menu.subItems && !isSidebarCompact && <ChevronDown size={16} className={`transition-transform ${isExpanded ? 'rotate-180' : ''}`} />}
                            </button>

                            {menu.subItems && isExpanded && !isSidebarCompact && (
                                <div className="ml-4 mt-1 space-y-1 border-l border-emerald-700/50 pl-2">
                                    {menu.subItems.map(sub => {
                                        const isSubDisabled = !isTahunSet && sub.id !== 'settings';
                                        return (
                                        <button
                                            key={sub.id}
                                            onClick={() => { if(!isSubDisabled) { setActiveMenu(sub.id); setIsSidebarOpen(false); } }}
                                            disabled={isSubDisabled}
                                            title={isSubDisabled ? 'Atur Tahun Ajaran Dulu' : sub.label}
                                            className={`w-full flex items-center gap-3 px-4 py-2 rounded-lg transition text-sm ${activeMenu === sub.id ? 'bg-emerald-600 text-white shadow' : 'text-emerald-200 hover:bg-emerald-700/50 hover:text-white'} ${isSubDisabled ? 'opacity-50 cursor-not-allowed' : ''}`}
                                        >
                                            <span className="font-medium">{sub.label}</span>
                                        </button>
                                    )})}
                                </div>
                            )}
                        </div>
                    );
                })}
            </div>
        </div>
        
        <div className="p-4 border-t border-emerald-700/50">
          <div className={`flex items-center ${isSidebarCompact ? 'justify-center' : 'gap-3'} ${isSidebarCompact ? 'px-0 py-3' : 'px-4 py-3'} bg-emerald-900/50 rounded-xl mb-4`}>
            <div className="w-10 h-10 rounded-full bg-emerald-600 flex items-center justify-center font-bold text-white">{currentUser?.name?.charAt(0) || 'U'}</div>
            {!isSidebarCompact && (
              <div className="flex-1 overflow-hidden">
                <p className="font-bold text-sm truncate text-white">{currentUser?.name}</p>
                <span className={`text-xs font-semibold uppercase px-2 py-0.5 rounded-full ${
                  currentUser?.role === 'admin' ? 'bg-emerald-500 text-white' :
                  currentUser?.role === 'guru' ? 'bg-blue-500 text-white' :
                  'bg-gray-500 text-white'
                }`}>{currentUser?.role || 'user'}</span>
              </div>
            )}
          </div>
          <button onClick={() => setIsLogoutModalOpen(true)} className={`w-full flex items-center justify-center ${isSidebarCompact ? 'px-0 py-3' : 'gap-2 px-4 py-2'} text-emerald-200 hover:text-white hover:bg-red-500/20 rounded-lg transition`} title="Keluar">
            <LogOut size={18} />
            {!isSidebarCompact && 'Keluar'}
          </button>
        </div>
      </div>
      
      <div className="flex-1 flex flex-col min-w-0 h-screen overflow-hidden">
        <header className="bg-white shadow-sm border-b px-6 py-4 flex items-center justify-between gap-4 print:hidden z-10 shrink-0">
          <div className="flex items-center gap-4">
            <button className="md:hidden text-gray-500 hover:text-emerald-600" onClick={() => setIsSidebarOpen(true)}><Menu size={24} /></button>
            <h1 className="text-xl font-bold text-gray-800 capitalize flex-1">{getMenuLabel()}</h1>
          </div>
          
          <div className="hidden sm:flex items-center gap-4">
              <AutoSaveIndicator />
              {isTahunSet ? (
                <div className="bg-emerald-100 text-emerald-800 px-3 py-1.5 rounded-lg text-sm font-bold border border-emerald-200 flex items-center gap-2">
                    <CheckCircle size={16} /> TA: {activeSetting.tahun} ({activeSetting.semester})
                </div>
              ) : (
                <div className="bg-red-100 text-red-800 px-3 py-1.5 rounded-lg text-sm font-bold border border-red-200 flex items-center gap-2 animate-pulse">
                    <AlertCircle size={16} /> Tahun Ajaran Belum Diatur
                </div>
              )}
              <CurrentTime />
          </div>
        </header>
        
        <main className="flex-1 overflow-hidden p-4 md:p-6 print:p-0 print:overflow-visible relative flex flex-col print:static">{renderContent()}</main>
      </div>

      <Modal isOpen={isLogoutModalOpen} onClose={() => setIsLogoutModalOpen(false)} title="Konfirmasi Keluar">
        <div className="space-y-4">
            <p className="text-gray-700">Apakah Anda yakin ingin keluar dari aplikasi?</p>
            <div className="flex justify-end gap-3 pt-4 border-t">
                <button onClick={() => setIsLogoutModalOpen(false)} className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded-lg font-medium transition">Batal</button>
                <button onClick={() => { setCurrentUser(null); setIsLogoutModalOpen(false); }} className="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg font-medium transition">Ya, Keluar</button>
            </div>
        </div>
      </Modal>
    </div>
  );
};

export default function App() {
  return <AppProvider><AppContext.Consumer>{({ currentUser }) => currentUser ? <Dashboard /> : <Login />}</AppContext.Consumer></AppProvider>;
}