/* eslint-disable security/detect-object-injection */
/* eslint-disable i18next/no-literal-string */
import React, { useState, useEffect, useRef, useMemo, createContext, useContext } from 'react';
import { 
  Menu, X, Home, Users, BookOpen, Settings, LayoutTemplate, 
  Printer, CheckSquare, LogOut, Plus, Trash2, Edit2, Save,
  Download, Upload, Share2, AlertCircle, CheckCircle, GripHorizontal,
  Type, User, CreditCard, Image as ImageIcon, Ruler, Type as TypeIcon, FileText,
  Columns, FileSignature, TrendingUp, UserX, Clock, Activity, ChevronDown,
  ZoomIn, ZoomOut, Maximize, Minimize, ChevronUp, Lock, Database
} from 'lucide-react';
import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';

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
  const SEMESTER_SPECIFIC_COLS = ['students', 'subjects', 'teachers', 'extracurriculars', 'presences', 'characterTraits'];

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
        if (!item.tahun) return true; // Legacy data: tampilkan di semua semester
        return item.tahun === activeSetting.tahun && item.semester === activeSetting.semester;
      });
    }
    return filtered;
  }, [allData, activeSetting]);

  const [loading, setLoading] = useState(true);
  const [notification, setNotification] = useState(null);

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
      if(!silent) showNotification('Gagal menyimpan data.', 'error');
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
    <AppContext.Provider value={{ data, allData, activeSetting, SEMESTER_SPECIFIC_COLS, currentUser, setCurrentUser, saveToDb, deleteFromDb, showNotification, addLog }}>
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
        let mostAbsent = { name: '-', total: 0, kelas: '-', detail: '' };

        if (!activeSetting.tahun) return { missingGrades, topStudent, mostAbsent };

        const subjectIds = data.subjects.map(s => s.id);

        data.classes.forEach(c => {
            const gradeDocId = getGradeDocId(c.id, data.classes, activeSetting, data.grades);
            const classGrades = data.grades.find(g => g.id === gradeDocId)?.data || {};
            const studentsInClass = getStudentsInClass(data.students, data.classes, c.id);
            const className = getClassNameFromValue(data.classes, c.id);

            if (studentsInClass.length > 0) {
                data.subjects.forEach(sub => {
                    let hasGrade = false;
                    studentsInClass.forEach(st => {
                        if (classGrades[st.id] && classGrades[st.id][sub.id] !== undefined && classGrades[st.id][sub.id] !== '') {
                            hasGrade = true;
                        }
                    });
                    if (!hasGrade) {
                        missingGrades.push({ subject: sub.nameId, guru: sub.guru || '-', kelas: className });
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
            }
        });

        return { missingGrades, topStudent, mostAbsent };
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
    
    const snapshotId = activeSetting ? `${(activeSetting.tahun || '').replace(/\//g, '-')}_${activeSetting.semester || '1'}` : null;
    const currentSnapshot = data.studentSnapshots.find(s => s.id === snapshotId);
    
    const handleSaveSnapshot = async () => {
        if (!activeSetting || !activeSetting.tahun) {
            showNotification('Aktifkan tahun ajaran terlebih dahulu!', 'error');
            return;
        }
        if (!confirm(`Simpan/Update Snapshot Santri untuk Tahun ${activeSetting.tahun} Semester ${activeSetting.semester}?`)) return;
        
        setIsProcessing(true);
        try {
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
    };

    const handleBackupExcel = () => {
        if (!activeSetting || !activeSetting.tahun) {
            showNotification('Aktifkan Tahun Ajaran terlebih dahulu untuk mem-backup data!', 'error');
            return;
        }

        try {
            const wb = XLSX.utils.book_new();
            const collections = ['settings', 'users', 'subjectCategories', 'masterSubjects', 'subjects', 'classes', 'students', 'teachers', 'grades', 'layouts', 'fonts', 'studentFields', 'presences', 'extracurriculars', 'characterTraits', 'studentSnapshots'];
            
            collections.forEach(col => {
                if (data[col] && data[col].length > 0) {
                    let filteredData = data[col];
                    // Hanya backup data yang sesuai dengan Tahun Ajaran yang aktif untuk koleksi transaksi/tahunan
                    if (['settings', 'grades', 'studentSnapshots'].includes(col)) {
                        filteredData = data[col].filter(item => item.tahun === activeSetting.tahun);
                    }
                    
                    if (filteredData.length === 0) return;

                    const flatData = filteredData.map(item => {
                        let newItem = {};
                        for (const key in item) {
                            if (typeof item[key] === 'object' && item[key] !== null) {
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
            });

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
        } catch (e) {
            console.error(e);
            showNotification('Gagal membuat file Excel', 'error');
        }
    };

    const handleRestoreExcel = (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = async (event) => {
            try {
                const fileData = new Uint8Array(event.target.result);
                const workbook = XLSX.read(fileData, {type: 'array'});
                
                let importedData = {};
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const rawData = XLSX.utils.sheet_to_json(worksheet);
                    importedData[sheetName] = rawData.map(item => {
                        let parsedItem = {};
                        for (const key in item) {
                            if (typeof item[key] === 'string' && (item[key].startsWith('{') || item[key].startsWith('['))) {
                                try { parsedItem[key] = JSON.parse(item[key]); }
                                catch (e) { parsedItem[key] = item[key]; }
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
                
                for (let tahun of importedTahunSet) {
                    if (currentTahunSet.has(tahun)) {
                        throw new Error(`Tahun Ajaran "${tahun}" dari file backup sudah ada di sistem. Restore dibatalkan untuk mencegah tertimpanya data.`);
                    }
                }

                if (!confirm('Tahun ajaran dari backup belum ada di sistem. Anda yakin ingin merestore data ini? (Aman: Data hanya akan ditambahkan)')) {
                    e.target.value = '';
                    return;
                }
                
                setIsProcessing(true);
                showNotification('Memulai proses restore data, mohon tunggu...', 'info');
                
                const collections = ['settings', 'users', 'subjectCategories', 'masterSubjects', 'subjects', 'classes', 'students', 'teachers', 'grades', 'layouts', 'fonts', 'studentFields', 'presences', 'extracurriculars', 'characterTraits', 'studentSnapshots'];
                
                let count = 0;
                for (const col of collections) {
                    if (Array.isArray(importedData[col])) {
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
                showNotification(`Restore selesai (${count} data dipulihkan)! Halaman akan dimuat ulang.`);
                setTimeout(() => window.location.reload(), 2000);
            } catch (err) {
                console.error(err);
                showNotification('Gagal memulihkan backup: ' + err.message, 'error');
                setIsProcessing(false);
            }
            e.target.value = '';
        };
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
        </div>
    );
};

const MasterData = ({ activeTab }) => {
  const { data, saveToDb, deleteFromDb, showNotification } = useContext(AppContext);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingItem, setEditingItem] = useState(null);
  const [formData, setFormData] = useState({});
  const [isAutoSaving, setIsAutoSaving] = useState(false);
    const [autoSaveEnabled, setAutoSaveEnabled] = useState(false);
    const [hideInactive, setHideInactive] = useState(false);

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
                  aValue = getSubjectClassLabel(a, data.classes);
                  bValue = getSubjectClassLabel(b, data.classes);
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
          const kelasLabel = getSubjectClassLabel(sub, data.classes);
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
        try {
            const dataBuf = await file.arrayBuffer();
            const workbook = XLSX.read(dataBuf, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
            let count = 0;
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const item = {};
                // Normalisasi kunci umum ke field yang dipakai aplikasi
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
                // Simpan (silent)
                // await memastikan urutan pengiriman ke DB sehingga tidak melebihi batas antrian
                // namun kita jalankan secara serial untuk kesederhanaan
                // Jika ada banyak data, pengguna disarankan memecah file.
                // eslint-disable-next-line no-await-in-loop
                await saveToDb(type, item.id, item, true);
                count++;
            }
            showNotification(`${count} data berhasil diimpor dari Excel!`);
        } catch (err) {
            console.error(err);
            showNotification('Gagal memproses file Excel. Pastikan format benar.', 'error');
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
                          <td className="p-3 text-center font-semibold text-gray-800">{getSubjectClassLabel(sub, data.classes) || 'Semua'}</td>
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
                        <div>
                            <div className="mb-4 flex gap-2">
                                <button onClick={() => generateExcelTemplate('students')} className="bg-emerald-600 text-white px-4 py-2 rounded-lg flex items-center gap-2 hover:bg-emerald-700 shadow-sm"><Download size={16}/> Download Template Excel</button>
                                <label className="bg-emerald-100 text-emerald-700 px-4 py-2 rounded-lg cursor-pointer flex items-center gap-2 hover:bg-emerald-200">
                                    <Upload size={18} /> Impor Excel <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => handleImportExcel(e, 'students')} />
                                </label>
                            </div>
                            <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                    <SortableHeader label="NIS" sortKey="nis" />
                    <SortableHeader label="Nama Santri" sortKey="nama" />
                    <SortableHeader label="Nama Arab" sortKey="nama_arab" />
                    <SortableHeader label="Kelas" sortKey="kelas" />
                    <th className="p-3 border-b text-center">Aksi</th>
                </tr></thead>
                <tbody>{sortedData.map(st => (<tr key={st.id} className="border-b hover:bg-gray-50"><td className="p-3">{st.nis}</td><td className="p-3 font-semibold">{st.nama}</td><td className="p-3 font-arabic" dir="rtl">{st.nama_arab}</td><td className="p-3">{getClassNameFromValue(data.classes, st.kelas)}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(st)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('students', st.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>))}</tbody>
              </table>
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

const renderDynamicTable = (el, data, studentGrades, classAverages = {}, isKatrol = false, mode = 'raport') => {
    const columns = el.columns || [];
    if(columns.length === 0) return <div className="p-4 border bg-red-50 text-red-500 text-xs">Tabel belum dikonfigurasi. Silakan edit kolom di panel kiri.</div>;

    const renderHeaders = () => (
        <thead>
            <tr className="bg-gray-100">
                {columns.map((col, idx) => (
                    <th key={idx} className="border border-black p-1 text-center font-bold" style={{width: `${col.width}%`}}>
                        {col.header}
                    </th>
                ))}
            </tr>
        </thead>
    );

    const renderRowCells = (sub, idx) => {
        return columns.map((col, cIdx) => {
            let content = '-';
            let style = {};
            
            let rawGrade = Number(studentGrades[sub.id]) || 0;
            let finalGrade = isKatrol ? Math.max(rawGrade, Number(sub.kkm || 0)) : rawGrade;
            let isRed = !isKatrol && rawGrade > 0 && rawGrade < Number(sub.kkm || 0);

            switch(col.type) {
                case 'NO': content = idx + 1; style={textAlign: 'center'}; break;
                case 'MAPEL_ID': content = sub.nameId || sub.name; break;
                case 'MAPEL_AR': content = sub.nameAr; style={textAlign: 'right', fontFamily: '"Amiri", "Scheherazade New", serif'}; break;
                case 'KKM': content = sub.kkm; style={textAlign: 'center'}; break;
                case 'NILAI': 
                    content = finalGrade || '-'; 
                    style={textAlign: 'center', fontWeight: 'bold', color: isRed ? 'red' : 'inherit'}; 
                    break;
                case 'RATA_KELAS': 
                    content = classAverages[sub.id] || '-'; 
                    style={textAlign: 'center'}; 
                    break;
                default: 
                    if(col.type.startsWith('PRESENCE_')) {
                        const pId = col.type.replace('PRESENCE_', '');
                        content = studentGrades[pId] || '-';
                        style={textAlign: 'center'};
                    } else if(col.type.startsWith('SIKAP_')) {
                        const sId = col.type.replace('SIKAP_', '');
                        content = studentGrades[sId] || '-';
                        style={textAlign: 'center'};
                    } else if(col.type.startsWith('EKSKUL_')) {
                        const eId = col.type.replace('EKSKUL_', '');
                        content = studentGrades[eId] || '-';
                        style={textAlign: 'center'};
                    }
            }
            return <td key={cIdx} className="border border-black p-1" style={style}>{content}</td>
        });
    };

    const subjectsToRender = mode === 'ijazah' ? data.subjects.filter(s => s.isIjazah) : data.subjects;

    if (el.groupByCategory) {
        const grouped = groupBy(subjectsToRender, 'kategori');
        let globalIndex = 0;
        return (
            <table className="w-full border-collapse border border-black text-sm" style={{ width: '100%', fontSize: `${el.fontSize}px`, fontFamily: el.fontFamily }}>
                {renderHeaders()}
                <tbody>
                    {Object.entries(grouped).sort(([a],[b]) => a.localeCompare(b)).map(([cat, subs]) => (
                        <React.Fragment key={cat}>
                            {cat && <tr><td colSpan={columns.length} className="border border-black p-1 font-bold bg-gray-50">{cat}</td></tr>}
                            {subs.map(sub => {
                                globalIndex++;
                                return <tr key={sub.id}>{renderRowCells(sub, globalIndex - 1)}</tr>
                            })}
                        </React.Fragment>
                    ))}
                </tbody>
            </table>
        );
    } else {
        return (
            <table className="w-full border-collapse border border-black text-sm" style={{ width: '100%', fontSize: `${el.fontSize}px`, fontFamily: el.fontFamily }}>
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
const defaultTableColumns = [
    { id: 'c1', header: 'No', type: 'NO', width: 5 },
    { id: 'c2', header: 'Mata Pelajaran', type: 'MAPEL_ID', width: 35 },
    { id: 'c3', header: 'المادة', type: 'MAPEL_AR', width: 35 },
    { id: 'c4', header: 'KKM', type: 'KKM', width: 10 },
    { id: 'c5', header: 'Nilai', type: 'NILAI', width: 15 }
];

const LayoutBuilder = () => {
    const { data, saveToDb, deleteFromDb, showNotification } = useContext(AppContext);
    const [activeLayout, setActiveLayout] = useState('raport');
    const [elements, setElements] = useState([]);
    const [newLayoutName, setNewLayoutName] = useState('');
    const [showNewLayoutForm, setShowNewLayoutForm] = useState(false);
    const [pageSize, setPageSize] = useState('A4');
    const [guides, setGuides] = useState({ h: [], v: [] });
    const [selectedElementId, setSelectedElementId] = useState(null);
    const canvasRef = useRef(null);

    const canvasWidth = pageDimensions[pageSize].width;
    const canvasHeight = pageDimensions[pageSize].height;
    
    const allFonts = useMemo(() => [...defaultFontOptions, ...(data.fonts || [])], [data.fonts]);

    useEffect(() => {
        const styleId = 'custom-fonts-style';
        let styleTag = document.getElementById(styleId);
        if (!styleTag) { styleTag = document.createElement('style'); styleTag.id = styleId; document.head.appendChild(styleTag); }
        const imports = data.fonts?.filter(f => f.url).map(f => `@import url('${f.url}');`).join('\n') || '';
        styleTag.innerHTML = imports;
    }, [data.fonts]);

    useEffect(() => {
        const savedLayout = data.layouts.find(l => l.id === activeLayout);
        if (savedLayout) {
            setElements(savedLayout.elements || []);
            setPageSize(savedLayout.pageSize || 'A4');
            setGuides(savedLayout.guides || { h: [], v: [] });
        } else {
            setElements([]); setPageSize('A4'); setGuides({ h: [], v: [] });
        }
        setSelectedElementId(null);
    }, [activeLayout, data.layouts]);

    const addElement = (type, customKey = null) => {
        let defaultContent = type === 'text' ? 'Teks Baru' : type === 'image' ? 'https://via.placeholder.com/150' : `{{${type}}}`;
        if (customKey) defaultContent = `{{${customKey}}}`;

        const newEl = {
            id: Date.now().toString(),
            type, content: defaultContent,
            x: 50, y: 50, fontSize: 14, fontFamily: 'Arial, sans-serif', fontWeight: 'normal',
            width: type === 'table_grades' ? 650 : type === 'image' ? 100 : 200,
            height: type === 'table_grades' ? 300 : type === 'image' ? 100 : 30,
            ...(type === 'table_grades' ? { columns: [...defaultTableColumns], groupByCategory: false } : {})
        };
        setElements([...elements, newEl]);
        setSelectedElementId(newEl.id);
    };

    const updateElement = (id, changes) => setElements(elements.map(el => el.id === id ? { ...el, ...changes } : el));
    const removeElement = (id) => { setElements(elements.filter(el => el.id !== id)); setSelectedElementId(null); };
    const saveLayout = () => saveToDb('layouts', activeLayout, { name: activeLayout, elements, pageSize, guides }, false, `Menyimpan desain layout ${activeLayout}`);

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
        saveToDb('layouts', layoutId, { name: newLayoutName, elements: [], pageSize: 'A4', guides: { v: [], h: [] } }, false, `Membuat layout baru: ${newLayoutName}`);
        setNewLayoutName('');
        setShowNewLayoutForm(false);
        setActiveLayout(layoutId);
    };

    const deleteLayout = (layoutId) => {
        if (confirm(`Hapus layout "${data.layouts.find(l => l.id === layoutId)?.name || layoutId}"?`)) {
            deleteFromDb('layouts', layoutId, false, `Menghapus layout ${layoutId}`);
            if (activeLayout === layoutId) {
                setActiveLayout(data.layouts.find(l => l.id !== layoutId)?.id || 'raport');
            }
        }
    };

    // Drag Logic
    const [draggingType, setDraggingType] = useState(null);
    const [dragIndex, setDragIndex] = useState(null);
    const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 });

    const handleElementMouseDown = (e, el) => {
        e.stopPropagation(); setSelectedElementId(el.id); setDraggingType('element'); setDragIndex(el.id);
        const rect = e.target.getBoundingClientRect();
        setDragOffset({ x: e.clientX - rect.left, y: e.clientY - rect.top });
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
            updateElement(dragIndex, { x: newX, y: newY });
        } else if (draggingType === 'guide_v') {
            const newGuides = { ...guides }; newGuides.v[dragIndex] = rawX; setGuides(newGuides);
        } else if (draggingType === 'guide_h') {
            const newGuides = { ...guides }; newGuides.h[dragIndex] = rawY; setGuides(newGuides);
        }
    };

    const handleMouseUp = () => {
        if (draggingType?.startsWith('guide_')) {
            const newGuides = { ...guides };
            if (draggingType === 'guide_v' && (guides.v[dragIndex] < -20 || guides.v[dragIndex] > canvasWidth + 20)) newGuides.v.splice(dragIndex, 1);
            if (draggingType === 'guide_h' && (guides.h[dragIndex] < -20 || guides.h[dragIndex] > canvasHeight + 20)) newGuides.h.splice(dragIndex, 1);
            setGuides(newGuides);
        }
        setDraggingType(null); setDragIndex(null);
    };

    const handleImageUpload = (e) => {
        const file = e.target.files[0];
        if (file) {
            if (file.size > 500000) return showNotification("Maksimal gambar 500kb.", "error");
            const reader = new FileReader();
            reader.onload = (ev) => updateElement(selectedElementId, { content: ev.target.result });
            reader.readAsDataURL(file);
        }
    };

    const activeEl = elements.find(e => e.id === selectedElementId);
    
    // Memberikan objek dummy default agar renderDynamicTable tidak crash saat proses desain layout
    const mockStudentGrades = {};
    const mockClassAverages = {};

    return (
        <div className="flex flex-col md:flex-row gap-6 h-[80vh]">
            <div className="w-full md:w-[320px] bg-white rounded-xl shadow-sm flex flex-col border border-gray-100 shrink-0 overflow-hidden">
                <div className="p-4 border-b bg-gray-50 flex items-center justify-between z-10 shrink-0">
                    <h3 className="font-bold text-gray-800 text-lg">Layout Builder</h3>
                </div>
                
                <div className="flex-1 overflow-y-auto p-4 space-y-4 custom-scrollbar">
                    <div className="flex gap-2">
                        <select className="flex-1 p-2 border rounded-lg bg-white text-sm font-bold text-emerald-800" value={activeLayout} onChange={e => setActiveLayout(e.target.value)}>
                            {data.layouts && data.layouts.map(l => <option key={l.id} value={l.id}>{l.name || l.id}</option>)}
                        </select>
                        <button onClick={() => setShowNewLayoutForm(!showNewLayoutForm)} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 rounded-lg text-sm font-bold transition" title="Tambah layout baru"><Plus size={16}/></button>
                        {data.layouts && data.layouts.length > 1 && <button onClick={() => deleteLayout(activeLayout)} className="bg-red-500 hover:bg-red-600 text-white px-3 rounded-lg text-sm font-bold transition" title="Hapus layout"><Trash2 size={16}/></button>}
                    </div>
                    <select className="w-full p-2 border rounded-lg bg-white text-sm font-bold text-blue-800" value={pageSize} onChange={e => setPageSize(e.target.value)}>
                        <option value="A4">Size: A4</option><option value="F4">Size: F4</option>
                    </select>
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
                        <button onClick={() => addElement('table_grades')} className="w-full bg-orange-50 hover:bg-orange-100 text-orange-700 py-2 rounded text-sm flex items-center justify-center gap-2"><Columns size={16}/> Tabel Nilai Dinamis</button>
                        
                        <p className="text-xs font-semibold text-gray-500 uppercase mt-4 mb-1">Variabel Santri & Wali</p>
                        <div className="grid grid-cols-2 gap-1">
                            <button onClick={() => addElement('nama_santri')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><User size={14}/> Nama</button>
                            <button onClick={() => addElement('nis')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><CreditCard size={14}/> NIS</button>
                            <button onClick={() => addElement('kelas')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><BookOpen size={14}/> Kelas</button>
                            <button onClick={() => addElement('catatan_wali')} className="col-span-2 bg-pink-50 hover:bg-pink-100 text-pink-700 py-1.5 rounded text-xs flex justify-center gap-1"><FileSignature size={14}/> Catatan Wali Kelas</button>
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
                    </div>

                    {selectedElementId && activeEl && (
                        <div className="space-y-3 pt-2 pb-8 bg-blue-50/30 p-3 rounded-lg border border-blue-100">
                            <p className="text-xs font-bold text-blue-800 uppercase flex items-center justify-between">
                                Sedang Edit: {activeEl.type === 'table_grades' ? 'Tabel' : activeEl.type === 'image' ? 'Gambar' : 'Teks'}
                                <button onClick={() => setSelectedElementId(null)} className="text-gray-400 hover:text-gray-700"><X size={14}/></button>
                            </p>
                            
                            {activeEl.type === 'text' && (
                                <textarea className="w-full p-2 border rounded text-sm focus:ring-2 outline-none min-h-[60px]" value={activeEl.content} onChange={e => updateElement(selectedElementId, { content: e.target.value })} />
                            )}

                            {activeEl.type === 'image' && (
                                <div className="bg-white p-2 rounded border space-y-2">
                                    <label className="block text-xs font-semibold text-gray-700">Upload Ulang Gambar (Maks 500kb):</label>
                                    <input type="file" accept="image/*" className="text-xs w-full" onChange={handleImageUpload} />
                                </div>
                            )}

                            <div className="flex gap-2">
                                <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Posisi X</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={Math.round(activeEl.x || 0)} onChange={e => updateElement(selectedElementId, { x: Number(e.target.value) })}/></div>
                                <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Posisi Y</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={Math.round(activeEl.y || 0)} onChange={e => updateElement(selectedElementId, { y: Number(e.target.value) })}/></div>
                            </div>

                            {activeEl.type !== 'table_grades' && activeEl.type !== 'image' && (
                                <>
                                    <div>
                                        <label className="text-[10px] text-gray-500 font-bold uppercase">Jenis Font</label>
                                        <select className="w-full p-1.5 border rounded text-sm outline-none bg-white" value={activeEl.fontFamily || 'Arial, sans-serif'} onChange={e => updateElement(selectedElementId, { fontFamily: e.target.value })}>
                                            {allFonts.map((font, idx) => <option key={idx} value={font.value}>{font.name}</option>)}
                                        </select>
                                    </div>
                                    <div className="flex gap-2">
                                        <div className="w-2/3"><label className="text-[10px] text-gray-500 font-bold uppercase">Ukuran Teks</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={activeEl.fontSize} onChange={e => updateElement(selectedElementId, { fontSize: Number(e.target.value) })}/></div>
                                        <div className="w-1/3 flex items-end"><button onClick={() => updateElement(selectedElementId, { fontWeight: activeEl.fontWeight === 'bold' ? 'normal' : 'bold' })} className={`w-full border p-1.5 rounded text-sm font-bold transition ${activeEl.fontWeight === 'bold' ? 'bg-gray-800 text-white' : 'bg-white hover:bg-gray-100'}`}>B</button></div>
                                    </div>
                                </>
                            )}

                            {(activeEl.type === 'image' || activeEl.type === 'table_grades') && (
                                <div className="flex gap-2">
                                    <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Lebar (Width)</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={activeEl.width} onChange={e => updateElement(selectedElementId, { width: Number(e.target.value) })}/></div>
                                    {activeEl.type === 'image' && (
                                        <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Tinggi (Height)</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={activeEl.height} onChange={e => updateElement(selectedElementId, { height: Number(e.target.value) })}/></div>
                                    )}
                                </div>
                            )}

                            {activeEl.type === 'table_grades' && (
                                <div className="mt-4 border-t pt-3 space-y-3">
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
                                        <input type="checkbox" checked={activeEl.groupByCategory || false} onChange={e => updateElement(selectedElementId, { groupByCategory: e.target.checked })} />
                                        Kelompokkan per Kategori Pelajaran
                                    </label>

                                    <div className="space-y-2 max-h-[300px] overflow-y-auto pr-1 custom-scrollbar">
                                        {(activeEl.columns || []).map((col, idx) => (
                                            <div key={col.id} className="bg-white p-2 border rounded border-l-4 border-l-orange-400 relative group">
                                                <button onClick={() => {
                                                    const newCols = [...activeEl.columns]; newCols.splice(idx, 1);
                                                    updateElement(selectedElementId, { columns: newCols });
                                                }} className="absolute top-1 right-1 text-red-400 hover:text-red-600 bg-white rounded-full"><X size={14}/></button>
                                                
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
                                                            <option value="MAPEL_ID">Nama Pelajaran (Indo)</option>
                                                            <option value="MAPEL_AR">Nama Pelajaran (Arab)</option>
                                                            <option value="KKM">Nilai KKM</option>
                                                            <option value="NILAI">Nilai Angka Santri</option>
                                                            <option value="RATA_KELAS">Rata-rata Kelas</option>
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
                                                    <div className="w-1/3 relative">
                                                        <input type="number" className="w-full text-[10px] p-1 border rounded pr-4" value={col.width} onChange={e => {
                                                            const newCols = [...activeEl.columns]; newCols[idx].width = Number(e.target.value);
                                                            updateElement(selectedElementId, { columns: newCols });
                                                        }} title="Lebar (%)"/>
                                                        <span className="absolute right-1.5 top-1 text-[10px] text-gray-400 pointer-events-none">%</span>
                                                    </div>
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                            )}

                            <button onClick={() => removeElement(selectedElementId)} className="w-full bg-red-50 hover:bg-red-100 text-red-600 py-2 rounded text-sm font-bold flex justify-center items-center gap-2 mt-4 transition"><Trash2 size={16}/> Hapus Elemen</button>
                        </div>
                    )}
                </div>
                
                <div className="p-4 border-t bg-gray-50 space-y-2 shrink-0 z-10">
                    <p className="text-[10px] text-gray-500 text-center leading-tight">Tarik garis gelap (Atas/Kiri kanvas) untuk Garis Bantu.</p>
                    <button onClick={saveLayout} className="w-full bg-emerald-600 hover:bg-emerald-700 text-white py-3 rounded-lg font-bold flex items-center justify-center gap-2 transition"><Save size={18}/> Simpan Layout</button>
                </div>
            </div>

            <div className="flex-1 bg-gray-200 rounded-xl overflow-auto p-4 flex justify-center items-start border border-gray-300 relative select-none custom-scrollbar">
                <div className="relative" style={{ width: `${canvasWidth}px`, height: `${canvasHeight}px`, transform: 'scale(0.8)', transformOrigin: 'top center', marginTop: '20px', marginLeft: '20px' }} onMouseMove={handleMouseMove} onMouseUp={handleMouseUp} onMouseLeave={handleMouseUp}>
                    <div onMouseDown={createHGuide} className="absolute top-[-20px] left-0 right-0 h-[20px] bg-slate-800 text-slate-300 text-xs flex justify-center items-center cursor-row-resize shadow-md rounded-t-md hover:bg-slate-700 transition"><Ruler size={12} className="mr-2"/> Tarik Garis Horizontal</div>
                    <div onMouseDown={createVGuide} className="absolute left-[-20px] top-0 bottom-0 w-[20px] bg-slate-800 text-slate-300 text-xs flex flex-col justify-center items-center cursor-col-resize shadow-md rounded-l-md hover:bg-slate-700 transition" style={{ writingMode: 'vertical-rl', textOrientation: 'mixed' }}>Tarik Vertikal <Ruler size={12} className="mt-2"/></div>

                    <div ref={canvasRef} className="absolute inset-0 bg-white shadow-xl overflow-hidden">
                        <div className="absolute inset-0 pointer-events-none" style={{ backgroundImage: 'linear-gradient(#f0f0f0 1px, transparent 1px), linear-gradient(90deg, #f0f0f0 1px, transparent 1px)', backgroundSize: '20px 20px', opacity: 0.5 }}></div>
                        
                        {guides.v.map((gx, i) => (<div key={`v-${i}`} onMouseDown={(e) => startDragGuide(e, 'v', i)} style={{ position: 'absolute', left: `${gx}px`, top: 0, bottom: 0, borderLeft: '1px dashed #0ea5e9', cursor: 'col-resize', zIndex: 10 }} className="hover:border-l-2 hover:border-blue-500 group"><div className="absolute -top-5 -left-4 bg-blue-500 text-white text-[10px] px-1 rounded opacity-0 group-hover:opacity-100">{Math.round(gx)}</div></div>))}
                        {guides.h.map((gy, i) => (<div key={`h-${i}`} onMouseDown={(e) => startDragGuide(e, 'h', i)} style={{ position: 'absolute', top: `${gy}px`, left: 0, right: 0, borderTop: '1px dashed #0ea5e9', cursor: 'row-resize', zIndex: 10 }} className="hover:border-t-2 hover:border-blue-500 group"><div className="absolute -left-6 -top-2 bg-blue-500 text-white text-[10px] px-1 rounded opacity-0 group-hover:opacity-100">{Math.round(gy)}</div></div>))}

                        {elements.map(el => {
                            const isSelected = selectedElementId === el.id;
                            const isDraggingThis = draggingType === 'element' && dragIndex === el.id;
                            
                            return (
                                <div key={el.id} onMouseDown={(e) => handleElementMouseDown(e, el)}
                                    style={{
                                        position: 'absolute', left: `${el.x}px`, top: `${el.y}px`, fontSize: `${el.fontSize}px`, fontFamily: el.fontFamily || 'Arial, sans-serif', fontWeight: el.fontWeight,
                                        width: (el.type === 'table_grades' || el.type === 'image') ? `${el.width}px` : 'auto', height: el.type === 'image' ? `${el.height}px` : 'auto',
                                        cursor: isDraggingThis ? 'grabbing' : 'grab', outline: isSelected ? '2px dashed #059669' : 'none', padding: (el.type === 'image' || el.type === 'table_grades') ? '0' : '2px', zIndex: isSelected ? 20 : 1
                                    }}
                                    className={`hover:outline hover:outline-1 hover:outline-gray-400 ${el.type === 'table_grades' ? 'bg-white' : ''}`}
                                >
                                    {el.type === 'table_grades' ? renderDynamicTable(el, data, mockStudentGrades, mockClassAverages) 
                                    : el.type === 'image' ? <img src={el.content} style={{ width: '100%', height: '100%', objectFit: 'contain', pointerEvents: 'none' }} alt="elemen" />
                                    : <div style={{ whiteSpace: 'pre-wrap' }}>{el.content}</div>}
                                    
                                    {isSelected && <div className="absolute -top-3 -left-3 bg-emerald-600 text-white rounded-full p-1 shadow z-30"><GripHorizontal size={12} /></div>}
                                </div>
                            );
                        })}
                    </div>
                </div>
            </div>
            
            <style>{`@import url('https://fonts.googleapis.com/css2?family=Amiri:ital,wght@0,400;0,700;1,400;1,700&display=swap');`}</style>
            <style>{`.custom-scrollbar::-webkit-scrollbar { width: 6px; } .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }`}</style>
        </div>
    );
};

// ==========================================
// ==========================================
// UTILITY: EXCEL IMPORT/EXPORT FOR GRADES
// ==========================================
const generateGradesExcelTemplate = (studentsInClass, subjectsInClass, className) => {
    const headers = ['No', 'NIS', 'Nama Santri'];
    subjectsInClass.forEach(sub => {
        headers.push(`${sub.nameId} - UTS`);
        headers.push(`${sub.nameId} - UAS`);
    });
    
    const rows = [headers];
    studentsInClass.forEach((st, idx) => {
        const row = [idx + 1, st.nis || '', st.nama];
        subjectsInClass.forEach(() => {
            row.push('');
            row.push('');
        });
        rows.push(row);
    });
    
    return rows;
};

const exportGradesToExcel = (grades, studentsInClass, subjectsInClass, className) => {
    const rows = generateGradesExcelTemplate(studentsInClass, subjectsInClass, className);
    
    // Fill in existing grades
    studentsInClass.forEach((st, idx) => {
        subjectsInClass.forEach((sub, subIdx) => {
            const uts = grades[st.id]?.[sub.id]?.uts || '';
            const uas = grades[st.id]?.[sub.id]?.uas || '';
            rows[idx + 1][3 + subIdx * 2] = uts;
            rows[idx + 1][3 + subIdx * 2 + 1] = uas;
        });
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const colWidths = [8, 15, 25];
    subjectsInClass.forEach(() => {
        colWidths.push(12, 12);
    });
    ws['!cols'] = colWidths.map(w => ({ wch: w }));
    
    // Freeze panes (first row + left 3 columns)
    ws['!freeze'] = { xSplit: 3, ySplit: 1 };
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Nilai');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `template_nilai_${className.replace(/\//g, '-')}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
};

const importGradesFromExcel = async (file, studentsInClass, subjectsInClass) => {
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
                    
                    const nis = row[nisKey];
                    const nama = row[namaKey];
                    
                    // Find matching student
                    const student = studentsInClass.find(st => 
                        (st.nis && st.nis.toString() === nis.toString()) ||
                        (st.nama && st.nama.toLowerCase().includes(nama.toString().toLowerCase()))
                    );
                    
                    if (!student) {
                        console.warn(`Siswa dengan NIS ${nis} atau nama ${nama} tidak ditemukan.`);
                        return;
                    }
                    
                    if (!importedGrades[student.id]) {
                        importedGrades[student.id] = {};
                    }
                    
                    // Parse each subject's UTS/UAS
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
    const { data, saveToDb, showNotification } = useContext(AppContext);
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
    const studentsInClass = getStudentsInClass(activeStudents, data.classes, selectedClass);
    const subjectsInClass = useMemo(() => sortSubjectsByCategory(filterSubjectsByClass(data.subjects, selectedClass, data.classes), data.subjectCategories), [data.subjects, data.subjectCategories, selectedClass, data.classes]);
    const gradeDocId = getGradeDocId(selectedClass, data.classes, activeSetting, data.grades);

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
        if (!selectedClass || subjectsInClass.length === 0) {
            showNotification('Pilih kelas dan pastikan ada mata pelajaran.', 'error');
            return;
        }
        const className = getClassNameFromValue(data.classes, selectedClass);
        exportGradesToExcel(localGrades, studentsInClass, subjectsInClass, className);
    };

    const handleImportGrades = async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        
        if (!selectedClass || subjectsInClass.length === 0) {
            showNotification('Pilih kelas dan pastikan ada mata pelajaran.', 'error');
            return;
        }

        try {
            setIsImporting(true);
            const importedGrades = await importGradesFromExcel(file, studentsInClass, subjectsInClass);
            
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
                                <th key={sub.id} className="p-2 border-b border-r border-emerald-500 text-center min-w-[120px] bg-emerald-700">
                                    <div className="font-bold truncate">{sub.nameId}</div>
                                    <div className="text-[10px] text-emerald-200 font-normal truncate mt-0.5">(Guru: {sub.guru || '-'})</div>
                                    <div className="text-[10px] text-yellow-300 font-bold truncate mt-0.5">KKM: {sub.kkm || '-'}</div>
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
                                                <div className="grid grid-cols-3 gap-2 items-center">
                                                    <input type="text" dir="auto" title="Nilai UTS (angka)" placeholder="UTS" className="w-full p-2 border rounded text-center font-semibold outline-none text-gray-800 focus:border-emerald-500" value={uts} onChange={e => handleGradeChange(st.id, sub.id, e.target.value, 'uts')} />
                                                    <input type="text" dir="auto" title="Nilai UAS (angka)" placeholder="UAS" className="w-full p-2 border rounded text-center font-semibold outline-none text-gray-800 focus:border-emerald-500" value={uas} onChange={e => handleGradeChange(st.id, sub.id, e.target.value, 'uas')} />
                                                    <div className={`rounded border p-2 text-sm font-bold text-center ${isRed ? 'text-red-600 bg-red-50 border-red-200' : 'text-gray-800 bg-gray-50 border-gray-200'}`}>
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
    const { data, addLog } = useContext(AppContext);
    const [selectedClass, setSelectedClass] = useState('');
    const [selectedStudent, setSelectedStudent] = useState('');
    const [useKatrol, setUseKatrol] = useState(false);
    
    const activeSetting = data.settings.find(s => s.isActive) || {};
    const activeStudents = getStudentsForYear(data.studentSnapshots, activeSetting, data.students);
    const studentsInClass = getStudentsInClass(activeStudents, data.classes, selectedClass);
    const studentData = activeStudents.find(s => s.id === selectedStudent);
    
    const layoutSettings = data.layouts.find(l => l.id === mode) || {};
    const activeLayout = layoutSettings.elements || [];
    const layoutPageSize = layoutSettings.pageSize || 'A4';
    const canvasWidth = pageDimensions[layoutPageSize].width;
    const canvasHeight = pageDimensions[layoutPageSize].height;

    useEffect(() => {
        const styleId = 'custom-fonts-style-print';
        let styleTag = document.getElementById(styleId);
        if (!styleTag) { styleTag = document.createElement('style'); styleTag.id = styleId; document.head.appendChild(styleTag); }
        const imports = data.fonts?.filter(f => f.url).map(f => `@import url('${f.url}');`).join('\n') || '';
        styleTag.innerHTML = imports;
    }, [data.fonts]);

    const gradeDocId = getGradeDocId(selectedClass, data.classes, activeSetting, data.grades);
    const classGradesDoc = data.grades.find(g => g.id === gradeDocId)?.data || {};
    const studentGrades = classGradesDoc[selectedStudent] || {};

    const classAverages = useMemo(() => {
        if(!gradeDocId) return {};
        const sums = {}; const counts = {};
        Object.values(classGradesDoc).forEach(sGrades => {
            Object.entries(sGrades).forEach(([k, v]) => {
                if(v !== '' && !isNaN(v)) { sums[k] = (sums[k]||0) + Number(v); counts[k] = (counts[k]||0) + 1; }
            });
        });
        const avgs = {};
        Object.keys(sums).forEach(k => avgs[k] = (sums[k]/counts[k]).toFixed(2));
        return avgs;
    }, [classGradesDoc, gradeDocId]);

    const handlePrint = () => { 
        document.title = "Cetak_Dokumen"; 
        addLog(`Mencetak ${mode} untuk ${studentData?.nama || 'Siswa'}`);
        window.print(); 
    };

    const handleSavePDF = () => {
        if(!studentData) return;
        const originalTitle = document.title;
        const ts = activeSetting.tahun ? activeSetting.tahun.replace(/\//g, '-') : 'tahun';
        const ss = activeSetting.semester || '1';
        const ns = studentData.nama.replace(/\s+/g, '_');
        const ks = getClassNameFromValue(data.classes, selectedClass).replace(/\s+/g, '_');
        document.title = mode === 'raport' ? `raport_${ts}_${ss}_${ns}_${ks}` : `ijazah_${ts}_${ns}_${ks}`;
        addLog(`Menyimpan ${mode} sebagai PDF untuk ${studentData.nama}`);
        window.print();
        setTimeout(() => { document.title = originalTitle; }, 2000);
    };

    const handleWA = () => {
        if(!studentData) return;
        const text = `Assalamu'alaikum. Berikut adalah pemberitahuan nilai ${mode} ananda *${studentData.nama}* kelas *${getClassNameFromValue(data.classes, selectedClass)}*. Harap hubungi sekolah untuk mengambil berkas cetak fisiknya.`;
        addLog(`Membagikan Info ${mode} via WA untuk ${studentData.nama}`);
        window.open(`https://wa.me/?text=${encodeURIComponent(text)}`, '_blank');
    };

    const getStyles = (el) => ({ position: 'absolute', left: `${el.x}px`, top: `${el.y}px`, fontSize: `${el.fontSize}px`, fontFamily: el.fontFamily || 'Arial, sans-serif', fontWeight: el.fontWeight, color: 'black' });

    const renderElement = (el) => {
        let content = el.content;
        if (studentData && typeof content === 'string') {
            content = content.replace('{{nama_santri}}', studentData.nama || '').replace('{{nis}}', studentData.nis || '').replace('{{kelas}}', getClassNameFromValue(data.classes, studentData.kelas) || '')
                             .replace('{{catatan_wali}}', studentGrades['catatan_wali'] || '');
            if (data.studentFields) data.studentFields.forEach(f => content = content.replace(new RegExp(`{{${f.key}}}`, 'g'), studentData[f.key] || ''));
        }

        if (el.type === 'table_grades') return <div style={{...getStyles(el), width: `${el.width}px`}}>{renderDynamicTable(el, data, studentGrades, classAverages, useKatrol, mode)}</div>;
        if (el.type === 'image') return <img src={el.content} style={{...getStyles(el), width: `${el.width}px`, height: `${el.height}px`, objectFit: 'contain'}} alt="C" />;
        return <div style={{...getStyles(el), whiteSpace: 'pre-wrap'}}>{content}</div>;
    };

    const cssPageSize = layoutPageSize === 'F4' ? '215.9mm 330.2mm' : 'A4';

    return (
        <div className="flex flex-col md:flex-row gap-6">
            <div className="w-full md:w-80 bg-white p-6 rounded-xl shadow-sm border border-gray-100 print:hidden shrink-0 h-fit">
                <h3 className="text-xl font-bold mb-4 capitalize">Cetak {mode}</h3>
                {!activeSetting.tahun && <div className="bg-yellow-50 text-yellow-800 p-3 rounded-lg text-sm mb-4 border border-yellow-200">Pastikan Admin mengaktifkan Tahun Ajaran di Master Data terlebih dahulu.</div>}
                <div className="space-y-4">
                    <select className="w-full p-2 border rounded-lg" value={selectedClass} onChange={e => {setSelectedClass(e.target.value); setSelectedStudent('');}}><option value="">-- Kelas --</option>{data.classes.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select>
                    <select className="w-full p-2 border rounded-lg" value={selectedStudent} onChange={e => setSelectedStudent(e.target.value)} disabled={!selectedClass}><option value="">-- Santri --</option>{studentsInClass.map(s => <option key={s.id} value={s.id}>{s.nama}</option>)}</select>
                    <div className="pt-4 border-t"><label className="flex items-center gap-2 bg-yellow-50 p-3 rounded-lg border border-yellow-200 cursor-pointer"><input type="checkbox" className="w-5 h-5 text-yellow-600" checked={useKatrol} onChange={e => setUseKatrol(e.target.checked)} /><div><p className="font-bold text-yellow-800 text-sm">Gunakan Nilai Katrol</p><p className="text-xs text-yellow-700">Nilai merah otomatis menjadi KKM</p></div></label></div>
                    <div className="pt-4 flex flex-col gap-3">
                        <button onClick={handlePrint} disabled={!selectedStudent} className="w-full bg-blue-600 hover:bg-blue-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition"><Printer size={18}/> Print Langsung</button>
                        <button onClick={handleSavePDF} disabled={!selectedStudent} className="w-full bg-red-600 hover:bg-red-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition"><Download size={18}/> Simpan sbg PDF</button>
                        <button onClick={handleWA} disabled={!selectedStudent} className="w-full bg-emerald-600 hover:bg-emerald-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition"><Share2 size={18}/> Kirim Info via WA</button>
                    </div>
                </div>
            </div>
            <div className="flex-1 bg-gray-200 p-8 rounded-xl overflow-auto flex justify-center border border-gray-300 print:bg-white print:p-0 print:border-none print:overflow-visible relative">
                {selectedStudent ? (
                    <div className="print-container bg-white shadow-xl relative print:shadow-none print:m-0" style={{ width: `${canvasWidth}px`, height: `${canvasHeight}px` }}>
                        {activeLayout.length > 0 ? activeLayout.map(el => <React.Fragment key={el.id}>{renderElement(el)}</React.Fragment>) : <div className="absolute inset-0 flex items-center justify-center text-gray-400 print:hidden">Layout belum disetting oleh Admin.</div>}
                    </div>
                ) : <div className="flex items-center justify-center h-full text-gray-400 print:hidden">Pilih santri untuk melihat preview {mode}.</div>}
            </div>
            <style>{`@import url('https://fonts.googleapis.com/css2?family=Amiri:ital,wght@0,400;0,700;1,400;1,700&display=swap'); @media print { body * { visibility: hidden; } .print-container, .print-container * { visibility: visible; } .print-container { position: absolute; left: 0; top: 0; margin: 0; padding: 0; box-shadow: none; } @page { size: ${cssPageSize}; margin: 0; } }`}</style>
        </div>
    );
};

// ==========================================
// LEGER KELAS
// ==========================================
const LeggerKelas = () => {
    const { data, addLog } = useContext(AppContext);
    const [selectedClass, setSelectedClass] = useState('');
    
    const activeSetting = data.settings.find(s => s.isActive) || {};
    const activeStudents = useMemo(() => getStudentsForYear(data.studentSnapshots, activeSetting, data.students), [data.studentSnapshots, activeSetting, data.students]);
    const students = useMemo(() => getStudentsInClass(activeStudents, data.classes, selectedClass), [activeStudents, data.classes, selectedClass]);
    const subjects = useMemo(() => sortSubjectsByCategory(filterSubjectsByClass(data.subjects, selectedClass, data.classes), data.subjectCategories), [data.subjects, data.subjectCategories, selectedClass, data.classes]);
    
    const gradeDocId = getGradeDocId(selectedClass, data.classes, activeSetting, data.grades);

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

    const exportExcel = () => {
        const className = getClassNameFromValue(data.classes, selectedClass);
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
                                <th rowSpan={3} className="p-3 border-b border-gray-700 text-center bg-gray-700">Total</th><th rowSpan={3} className="p-3 border-b border-gray-700 text-center bg-gray-700">Rata-rata</th><th rowSpan={3} className="p-3 border-b border-gray-700 text-center bg-gray-700">Predikat</th>
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
const Dashboard = () => {
  const { currentUser, setCurrentUser } = useContext(AppContext);
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
    { id: 'dashboard', label: 'Dashbord', icon: Home, roles: ['admin', 'user'] },
    { id: 'master_data', label: 'Master Data', icon: Users, roles: ['admin'], subItems: masterDataSubItems },
    { id: 'layout_builder', label: 'Desain Layout', icon: LayoutTemplate, roles: ['admin'] },
    { id: 'input_nilai', label: 'Input Nilai', icon: CheckSquare, roles: ['admin', 'user'], subItems: inputNilaiSubItems },
    { id: 'legger', label: 'Legger Kelas', icon: BookOpen, roles: ['admin', 'user'] },
    { id: 'cetak_raport', label: 'Cetak Raport', icon: Printer, roles: ['admin', 'user'] },
    { id: 'cetak_ijazah', label: 'Cetak Ijazah', icon: Printer, roles: ['admin', 'user'] },
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
      case 'cetak_raport': return <CetakDokumen mode="raport" />;
      case 'cetak_ijazah': return <CetakDokumen mode="ijazah" />;
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

                    return (
                        <div key={menu.id} className="mb-1">
                            <button 
                                onClick={() => { 
                                    if(menu.subItems) {
                                        setExpandedMenu(isExpanded ? '' : menu.id);
                                        if (!isExpanded && menu.subItems.length > 0) {
                                            setActiveMenu(menu.subItems[0].id);
                                        }
                                    } else {
                                        setActiveMenu(menu.id); setIsSidebarOpen(false); setExpandedMenu('');
                                    }
                                }} 
                                title={menu.label}
                                className={`w-full flex items-center ${isSidebarCompact ? 'justify-center px-2 py-3' : 'justify-between px-4 py-3'} rounded-xl transition ${isActive && !menu.subItems ? 'bg-emerald-600 text-white shadow-lg' : 'text-emerald-100 hover:bg-emerald-700/50'}`}
                            >
                                <div className={`flex items-center ${isSidebarCompact ? 'justify-center gap-0' : 'gap-3'}`}>
                                    <Icon size={20} />
                                    <span className={`${isSidebarCompact ? 'hidden' : 'font-medium'}`}>{menu.label}</span>
                                </div>
                                {menu.subItems && !isSidebarCompact && <ChevronDown size={16} className={`transition-transform ${isExpanded ? 'rotate-180' : ''}`} />}
                            </button>

                            {menu.subItems && isExpanded && !isSidebarCompact && (
                                <div className="ml-4 mt-1 space-y-1 border-l border-emerald-700/50 pl-2">
                                    {menu.subItems.map(sub => (
                                        <button
                                            key={sub.id}
                                            onClick={() => { setActiveMenu(sub.id); setIsSidebarOpen(false); }}
                                            className={`w-full flex items-center gap-3 px-4 py-2 rounded-lg transition text-sm ${activeMenu === sub.id ? 'bg-emerald-600 text-white shadow' : 'text-emerald-200 hover:bg-emerald-700/50 hover:text-white'}`}
                                        >
                                            <span className="font-medium">{sub.label}</span>
                                        </button>
                                    ))}
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
                <p className="text-xs text-emerald-300 uppercase">{currentUser?.role}</p>
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
          
          <div className="hidden sm:block">
              <CurrentTime />
          </div>
        </header>
        
        <main className="flex-1 overflow-hidden p-4 md:p-6 print:p-0 print:overflow-visible relative flex flex-col">{renderContent()}</main>
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
  useEffect(() => {
    let link = document.querySelector("link[rel~='icon']");
    if (!link) {
      link = document.createElement('link');
      link.rel = 'icon';
      document.head.appendChild(link);
    }
    link.href = 'https://i.ibb.co.com/DfZSFRsP/Chat-GPT-Image-3-Mei-2026-04-08-56.png';
    document.title = "Rapijaz - Ponpes Imam Syafi'i";
  }, []);

  return <AppProvider><AppContext.Consumer>{({ currentUser }) => currentUser ? <Dashboard /> : <Login />}</AppContext.Consumer></AppProvider>;
}