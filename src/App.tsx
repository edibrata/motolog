import React, { useState, useEffect, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  Wrench, RefreshCw, Grid, Plus, Edit2, Trash2, X, ChevronDown, Search, ArrowUp, Paperclip, Camera, AlertTriangle, Bell
} from 'lucide-react';
import { Vehicle, ServiceRecord, ToastMessage, RecordItem } from './types';
import { 
  CLOUD_DATABASE_URL, INTEL_CONFIG, formatCurrency, formatDate, exportToExcel, backupJSON 
} from './utils';
import { doc, setDoc, getDoc, onSnapshot } from 'firebase/firestore';
import { ref, uploadBytes, getDownloadURL } from 'firebase/storage';
import { db, storage } from './lib/firebase';
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid } from 'recharts';
import { useSwipeable } from 'react-swipeable';
import { motion, AnimatePresence } from 'motion/react';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

const THEMES = [
  'indigo',
  'emerald',
  'rose',
  'amber',
  'cyan',
  'violet',
  'fuchsia'
];

type ThemeClasses = {
  bgBase: string;
  bgLight: string;
  bgHover: string;
  text: string;
  textLight: string;
  border: string;
  borderFocus: string;
  ringFocus: string;
  shadow: string;
  shadowLg: string;
  shadowHover: string;
};

const THEME_MAP: Record<string, ThemeClasses> = {
  indigo: { bgBase: 'bg-indigo-600', bgLight: 'bg-indigo-50', bgHover: 'hover:bg-indigo-700', text: 'text-indigo-600', textLight: 'text-indigo-700', border: 'border-indigo-100', borderFocus: 'focus:border-indigo-500', ringFocus: 'focus:ring-indigo-50', shadow: 'shadow-indigo-500/50', shadowLg: 'shadow-indigo-600/30', shadowHover: 'shadow-indigo-600/20' },
  emerald: { bgBase: 'bg-emerald-600', bgLight: 'bg-emerald-50', bgHover: 'hover:bg-emerald-700', text: 'text-emerald-600', textLight: 'text-emerald-700', border: 'border-emerald-100', borderFocus: 'focus:border-emerald-500', ringFocus: 'focus:ring-emerald-50', shadow: 'shadow-emerald-500/50', shadowLg: 'shadow-emerald-600/30', shadowHover: 'shadow-emerald-600/20' },
  rose: { bgBase: 'bg-rose-600', bgLight: 'bg-rose-50', bgHover: 'hover:bg-rose-700', text: 'text-rose-600', textLight: 'text-rose-700', border: 'border-rose-100', borderFocus: 'focus:border-rose-500', ringFocus: 'focus:ring-rose-50', shadow: 'shadow-rose-500/50', shadowLg: 'shadow-rose-600/30', shadowHover: 'shadow-rose-600/20' },
  amber: { bgBase: 'bg-amber-600', bgLight: 'bg-amber-50', bgHover: 'hover:bg-amber-700', text: 'text-amber-600', textLight: 'text-amber-700', border: 'border-amber-100', borderFocus: 'focus:border-amber-500', ringFocus: 'focus:ring-amber-50', shadow: 'shadow-amber-500/50', shadowLg: 'shadow-amber-600/30', shadowHover: 'shadow-amber-600/20' },
  cyan: { bgBase: 'bg-cyan-600', bgLight: 'bg-cyan-50', bgHover: 'hover:bg-cyan-700', text: 'text-cyan-600', textLight: 'text-cyan-700', border: 'border-cyan-100', borderFocus: 'focus:border-cyan-500', ringFocus: 'focus:ring-cyan-50', shadow: 'shadow-cyan-500/50', shadowLg: 'shadow-cyan-600/30', shadowHover: 'shadow-cyan-600/20' },
  violet: { bgBase: 'bg-violet-600', bgLight: 'bg-violet-50', bgHover: 'hover:bg-violet-700', text: 'text-violet-600', textLight: 'text-violet-700', border: 'border-violet-100', borderFocus: 'focus:border-violet-500', ringFocus: 'focus:ring-violet-50', shadow: 'shadow-violet-500/50', shadowLg: 'shadow-violet-600/30', shadowHover: 'shadow-violet-600/20' },
  fuchsia: { bgBase: 'bg-fuchsia-600', bgLight: 'bg-fuchsia-50', bgHover: 'hover:bg-fuchsia-700', text: 'text-fuchsia-600', textLight: 'text-fuchsia-700', border: 'border-fuchsia-100', borderFocus: 'focus:border-fuchsia-500', ringFocus: 'focus:ring-fuchsia-50', shadow: 'shadow-fuchsia-500/50', shadowLg: 'shadow-fuchsia-600/30', shadowHover: 'shadow-fuchsia-600/20' },
};

const getThemeClasses = (color: string | undefined): ThemeClasses => {
  return THEME_MAP[color || 'indigo'] || THEME_MAP['indigo'];
};

export default function App() {
  const [vehicles, setVehicles] = useState<Vehicle[]>(() => {
    const saved = localStorage.getItem('motoVehicles');
    const parsed = saved ? JSON.parse(saved) : [{ id: 'default', name: 'Motor Utama', manualKM: 0, themeColor: 'indigo' }];
    return parsed.map((v: any, index: number) => ({
      ...v,
      themeColor: v.themeColor || THEMES[index % THEMES.length]
    }));
  });
  const [currentVehicleId, setCurrentVehicleId] = useState<string>(() => {
    return localStorage.getItem('motoCurrentVehicleId') || 'default';
  });
  const [records, setRecords] = useState<ServiceRecord[]>(() => {
    const saved = localStorage.getItem('motorRecords');
    return saved ? JSON.parse(saved) : [];
  });

  useEffect(() => {
    if (vehicles.length > 0 && !vehicles.some(v => v.id === currentVehicleId)) {
      setCurrentVehicleId(vehicles[0].id);
    }
  }, [vehicles, currentVehicleId]);
  
  const isRemoteUpdate = useRef(false);
  const isFirstMount = useRef(true);

  const TABS = ['home', 'analysis', 'history'] as const;
  const [currentMainTab, setCurrentMainTab] = useState<typeof TABS[number]>('home');
  const [tabDirection, setTabDirection] = useState(0);

  const handleTabChange = (newTab: typeof TABS[number]) => {
    if (newTab === currentMainTab) return;
    const currentIndex = TABS.indexOf(currentMainTab);
    const newIndex = TABS.indexOf(newTab);
    setTabDirection(newIndex > currentIndex ? 1 : -1);
    setCurrentMainTab(newTab);
  };

  const [historySearchQuery, setHistorySearchQuery] = useState('');
  const [historyMonthFilter, setHistoryMonthFilter] = useState('');
  const [historyPriceFilter, setHistoryPriceFilter] = useState('');
  
  const swipeHandlers = useSwipeable({
    onSwipedLeft: () => {
      const idx = TABS.indexOf(currentMainTab);
      if (idx < TABS.length - 1) handleTabChange(TABS[idx + 1]);
    },
    onSwipedRight: () => {
      const idx = TABS.indexOf(currentMainTab);
      if (idx > 0) handleTabChange(TABS[idx - 1]);
    },
    trackMouse: true
  });

  // Modals state
  const [isKMModalOpen, setIsKMModalOpen] = useState(false);
  const [isServiceModalOpen, setIsServiceModalOpen] = useState(false);
  const [isVehicleModalOpen, setIsVehicleModalOpen] = useState(false);
  const [isCloudModalOpen, setIsCloudModalOpen] = useState(false);
  const [activeImagePreview, setActiveImagePreview] = useState<string | null>(null);

  // Confirm Dialog state
  const [confirmDialog, setConfirmDialog] = useState({
    isOpen: false,
    title: '',
    message: '',
    onConfirm: () => {},
    onCancel: () => {},
  });

  const confirmAction = (title: string, message: string, onConfirm: () => void, onCancel: () => void = () => {}) => {
    setConfirmDialog({
      isOpen: true,
      title,
      message,
      onConfirm: () => {
        onConfirm();
        setConfirmDialog(prev => ({ ...prev, isOpen: false }));
      },
      onCancel: () => {
        onCancel();
        setConfirmDialog(prev => ({ ...prev, isOpen: false }));
      }
    });
  };

  // Form states
  const [manualKMInput, setManualKMInput] = useState('');
  
  const [editId, setEditId] = useState('');
  const [serviceDate, setServiceDate] = useState(() => new Date().toISOString().split('T')[0]);
  const [odometer, setOdometer] = useState('');
  const [workshop, setWorkshop] = useState('');
  const [serviceItems, setServiceItems] = useState<{id: string, name: string, price: string, type?: 'jasa' | 'part'}[]>([{
    id: Date.now().toString(), name: '', price: '', type: 'part'
  }]);
  const [attachmentFile, setAttachmentFile] = useState<File | null>(null);
  const [attachmentUrl, setAttachmentUrl] = useState<string>('');
  const [isUploading, setIsUploading] = useState(false);

  const [newVehicleName, setNewVehicleName] = useState('');
  const [editVehicleId, setEditVehicleId] = useState('');
  const [vehiclePhotoFile, setVehiclePhotoFile] = useState<File | null>(null);
  const [vehiclePhotoPreview, setVehiclePhotoPreview] = useState<string | null>(null);
  const [isVehicleUploading, setIsVehicleUploading] = useState(false);
  const vehiclePhotoInputRef = useRef<HTMLInputElement>(null);

  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const [isSyncing, setIsSyncing] = useState(false);
  
  // Reminder/Notification system states
  const [isReminderModalOpen, setIsReminderModalOpen] = useState(false);
  const [reminderThreshold, setReminderThreshold] = useState<number>(() => {
    return parseInt(localStorage.getItem('motoReminderThreshold') || '500');
  });
  const [notificationPermission, setNotificationPermission] = useState<string>(() => {
    if (typeof window !== 'undefined' && 'Notification' in window) {
      return Notification.permission;
    }
    return 'default';
  });

  const requestNotificationPermission = async () => {
    if (typeof window !== 'undefined' && 'Notification' in window) {
      try {
        const permission = await Notification.requestPermission();
        setNotificationPermission(permission);
        if (permission === 'granted') {
          showToast("Izin notifikasi browser diberikan!");
          new Notification("MotoJournal", {
            body: "Terima kasih! Notifikasi pengingat jadwal servis telah aktif.",
            icon: "/favicon.ico"
          });
        } else {
          showToast("Izin notifikasi ditolak", "error");
        }
      } catch (err) {
        showToast("Gagal meminta izin notifikasi", "error");
      }
    } else {
      showToast("Notifikasi browser tidak didukung di perangkat ini", "error");
    }
  };

  const checkAndNotifyService = (force: boolean = false) => {
    if (!activeVehicle) return;
    
    const overdueList: string[] = [];
    const approachingList: string[] = [];
    
    const currentKM = activeVehicle.manualKM || 0;
    
    INTEL_CONFIG.forEach(cfg => {
      const matches = vehicleRecords.filter(r => 
        r.items.some(item => cfg.kws.some(kw => item.name.toLowerCase().includes(kw)))
      );
      const lastRec = matches.length ? matches[0] : null;
      const lastKM = lastRec ? lastRec.odometer : 0;
      let progress = 0;
      if (lastRec) {
          progress = ((currentKM - lastKM) / cfg.interval) * 100;
      } else if (currentKM > 0) {
          progress = (currentKM / cfg.interval) * 100;
      }
      const nextKM = lastKM + cfg.interval;
      const remKM = nextKM - currentKM;
      
      if (remKM <= 0) {
        overdueList.push(cfg.name);
      } else if (remKM <= reminderThreshold) {
        approachingList.push(`${cfg.name} (${formatCurrency(remKM)} KM lagi)`);
      }
    });

    if (overdueList.length > 0 || approachingList.length > 0) {
      let bodyText = "";
      if (overdueList.length > 0) {
        bodyText += `Perlu servis segera: ${overdueList.join(', ')}. `;
      }
      if (approachingList.length > 0) {
        bodyText += `Mendekati jadwal: ${approachingList.join(', ')}.`;
      }

      // Show browser Notification
      if (typeof window !== 'undefined' && 'Notification' in window && Notification.permission === 'granted') {
        try {
          new Notification(`Pengingat Servis ${activeVehicle.name}`, {
            body: bodyText,
            icon: "/favicon.ico"
          });
        } catch (e) {
          console.error("Failed to send Notification:", e);
        }
      }

      // Show in-app toast
      showToast(`📢 ${overdueList.length ? 'Ada komponen kritis!' : 'Ada komponen mendekati jadwal servis.'} Cek tab Analisis.`, "error");
    } else {
      if (force) {
        showToast("✅ Semua komponen kendaraan dalam kondisi prima!");
      }
    }
  };

  useEffect(() => {
    localStorage.setItem('motoReminderThreshold', reminderThreshold.toString());
  }, [reminderThreshold]);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Real-time synchronization from Cloud
  useEffect(() => {
    const unsubscribe = onSnapshot(doc(db, 'motojournal', 'sync_data'), (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        if (data && Array.isArray(data.vehicles) && Array.isArray(data.records)) {
           const localV = localStorage.getItem('motoVehicles');
           const localR = localStorage.getItem('motorRecords');
           if (JSON.stringify(data.vehicles) !== localV || JSON.stringify(data.records) !== localR) {
              isRemoteUpdate.current = true;
              setVehicles(data.vehicles.map((v: any, idx: number) => ({
                 ...v,
                 themeColor: v.themeColor || THEMES[idx % THEMES.length]
              })));
              setRecords(data.records);
              showToast("Data disinkronkan otomatis dari Cloud");
           }
        }
      }
    });
    return () => unsubscribe();
  }, []);

  // Sync to Local Storage
  useEffect(() => {
    localStorage.setItem('motoVehicles', JSON.stringify(vehicles));
    localStorage.setItem('motorRecords', JSON.stringify(records));
    localStorage.setItem('motoCurrentVehicleId', currentVehicleId);
  }, [vehicles, records, currentVehicleId]);

  // Automatically to Cloud on local changes
  useEffect(() => {
    if (isFirstMount.current) {
      isFirstMount.current = false;
      return;
    }

    if (isRemoteUpdate.current) {
      isRemoteUpdate.current = false;
      return;
    }

    const syncToCloud = async () => {
      setIsSyncing(true);
      try {
        const docRef = doc(db, 'motojournal', 'sync_data');
        await setDoc(docRef, { vehicles, records });
      } catch (e) {
        console.error("Auto Sync Error:", e);
      } finally {
        setIsSyncing(false);
      }
    };
    
    const timeout = setTimeout(syncToCloud, 2000);
    return () => clearTimeout(timeout);
  }, [vehicles, records]);

  const showToast = (message: string, type: 'success' | 'error' = 'success') => {
    const id = Date.now().toString();
    setToasts(prev => [...prev, { id, message, type }]);
    setTimeout(() => {
      setToasts(prev => prev.filter(t => t.id !== id));
    }, 3000);
  };

  const activeVehicle = vehicles.find(v => v.id === currentVehicleId) || vehicles[0];
  const theme = getThemeClasses(activeVehicle?.themeColor);
  const vehicleRecords = useMemo(() => {
    if (!activeVehicle) return [];
    return records
      .filter(r => r.vehicleId === activeVehicle.id)
      .sort((a, b) => {
        const timeDiff = new Date(b.date).getTime() - new Date(a.date).getTime();
        if (timeDiff === 0) {
          return b.odometer - a.odometer; // sort by odometer descending if same date
        }
        return timeDiff;
      });
  }, [records, currentVehicleId]);

  const lastNotifiedRef = useRef<string>('');

  useEffect(() => {
    if (!activeVehicle) return;
    const key = `${activeVehicle.id}_${activeVehicle.manualKM}`;
    if (lastNotifiedRef.current === key) return;
    
    // Skip on very first mount
    if (isFirstMount.current) {
      return;
    }

    const timer = setTimeout(() => {
      checkAndNotifyService(false);
      lastNotifiedRef.current = key;
    }, 1500);
    return () => clearTimeout(timer);
  }, [activeVehicle?.manualKM, activeVehicle?.id, vehicleRecords]);
  
  const latestRecordByOdo = useMemo(() => {
    if (!vehicleRecords.length) return null;
    return vehicleRecords.reduce((max, r) => (r.odometer > max.odometer ? r : max), vehicleRecords[0]);
  }, [vehicleRecords]);

  const componentsNeedingService = useMemo(() => {
    if (!activeVehicle) return [];
    const currentKM = activeVehicle.manualKM || 0;
    return INTEL_CONFIG.filter(cfg => {
        const matches = vehicleRecords.filter(r => 
          r.items.some(item => cfg.kws.some(kw => item.name.toLowerCase().includes(kw)))
        );
        const lastRec = matches.length ? matches[0] : null;
        const lastKM = lastRec ? lastRec.odometer : 0;
        let progress = 0;
        if (lastRec) {
            progress = ((currentKM - lastKM) / cfg.interval) * 100;
        } else if (currentKM > 0) {
            progress = (currentKM / cfg.interval) * 100;
        }
        return progress >= 100;
    });
  }, [activeVehicle, vehicleRecords]);

  const componentsApproachingService = useMemo(() => {
    if (!activeVehicle) return [];
    const currentKM = activeVehicle.manualKM || 0;
    return INTEL_CONFIG.filter(cfg => {
        const matches = vehicleRecords.filter(r => 
          r.items.some(item => cfg.kws.some(kw => item.name.toLowerCase().includes(kw)))
        );
        const lastRec = matches.length ? matches[0] : null;
        const lastKM = lastRec ? lastRec.odometer : 0;
        let progress = 0;
        if (lastRec) {
            progress = ((currentKM - lastKM) / cfg.interval) * 100;
        } else if (currentKM > 0) {
            progress = (currentKM / cfg.interval) * 100;
        }
        const nextKM = lastKM + cfg.interval;
        const remKM = nextKM - currentKM;
        return remKM > 0 && remKM <= reminderThreshold;
    });
  }, [activeVehicle, vehicleRecords, reminderThreshold]);

  const filteredRecords = vehicleRecords.filter(record => {
    let match = true;
    if (historySearchQuery) {
        const q = historySearchQuery.toLowerCase();
        const workshopMatch = (record.workshop || '').toLowerCase().includes(q);
        const itemsMatch = record.items.some(i => i.name.toLowerCase().includes(q));
        if (!workshopMatch && !itemsMatch) match = false;
    }
    if (historyMonthFilter) {
        if (!record.date.startsWith(historyMonthFilter)) match = false;
    }
    if (historyPriceFilter) {
        const [min, max] = historyPriceFilter.split('-').map(Number);
        if (record.totalCost < min || (max && record.totalCost > max)) match = false;
    }
    return match;
  });

  const expenseData = useMemo(() => {
    const monthly: Record<string, number> = {};
    vehicleRecords.forEach(r => {
       const d = new Date(r.date);
       const monthYear = `${d.toLocaleString('default', { month: 'short' })} '${d.getFullYear().toString().substring(2)}`;
       if (!monthly[monthYear]) monthly[monthYear] = 0;
       monthly[monthYear] += r.totalCost;
    });
    return Object.entries(monthly).reverse().map(([name, cost]) => ({ name, cost }));
  }, [vehicleRecords]);

  const odometerData = useMemo(() => {
    const chronos = [...vehicleRecords].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    return chronos.map(r => {
        const d = new Date(r.date);
        return {
            date: `${d.getDate()} ${d.toLocaleString('default', { month: 'short' })} '${d.getFullYear().toString().substring(2)}`,
            km: r.odometer
        };
    });
  }, [vehicleRecords]);

  const dailyAverageKM = useMemo(() => {
    if (vehicleRecords.length < 2) return 15; // default 15 km/day
    const oldest = vehicleRecords[vehicleRecords.length - 1];
    const newest = vehicleRecords[0];
    const days = (new Date(newest.date).getTime() - new Date(oldest.date).getTime()) / (1000 * 3600 * 24);
    if (days === 0) return 15;
    return Math.max(1, (newest.odometer - oldest.odometer) / days);
  }, [vehicleRecords]);

  const saveManualKM = () => {
    const newKM = parseInt(manualKMInput) || 0;
    if (activeVehicle) {
      setVehicles(prev => prev.map(v => v.id === activeVehicle.id ? { ...v, manualKM: newKM } : v));
    }
    setIsKMModalOpen(false);
    showToast("Odometer diperbarui");
  };

  const handleAddServiceItem = () => {
    setServiceItems(prev => [...prev, { id: Date.now().toString(), name: '', price: '', type: 'part' }]);
  };

  const calculateFormTotal = () => {
    return serviceItems.reduce((acc, item) => acc + (parseInt(item.price) || 0), 0);
  };

  const handleServiceSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (isUploading) return;
    
    let uploadedUrl = attachmentUrl;
    
    if (attachmentFile) {
      setIsUploading(true);
      try {
        uploadedUrl = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (e) => {
            const img = new Image();
            img.onload = () => {
              const canvas = document.createElement('canvas');
              const MAX_WIDTH = 800;
              const MAX_HEIGHT = 800;
              let width = img.width;
              let height = img.height;

              if (width > height) {
                if (width > MAX_WIDTH) {
                  height *= MAX_WIDTH / width;
                  width = MAX_WIDTH;
                }
              } else {
                if (height > MAX_HEIGHT) {
                  width *= MAX_HEIGHT / height;
                  height = MAX_HEIGHT;
                }
              }
              canvas.width = width;
              canvas.height = height;
              const ctx = canvas.getContext('2d');
              ctx?.drawImage(img, 0, 0, width, height);
              resolve(canvas.toDataURL('image/jpeg', 0.75));
            };
            img.onerror = () => reject(new Error('Gagal memproses gambar'));
            img.src = e.target?.result as string;
          };
          reader.onerror = () => reject(new Error('Gagal membaca file'));
          reader.readAsDataURL(attachmentFile);
        });
      } catch (err) {
        showToast("Gagal memproses foto lampiran");
        setIsUploading(false);
        return;
      }
      setIsUploading(false);
    }
    
    const validItems = serviceItems.filter(i => i.name.trim() !== '').map(i => ({
      name: i.name.trim(),
      price: parseInt(i.price) || 0,
      type: i.type || 'part'
    }));
    
    if (!validItems.length) {
      showToast("Tolong masukkan minimal 1 komponen/jasa");
      return;
    }
    
    const odoValue = parseInt(odometer);
    const totalCost = validItems.reduce((acc, item) => acc + item.price, 0);
    
    const newRecord: ServiceRecord = {
      id: editId || Date.now().toString(),
      vehicleId: activeVehicle?.id || currentVehicleId,
      date: serviceDate,
      odometer: odoValue,
      workshop: workshop,
      items: validItems,
      totalCost,
      attachmentUrl: uploadedUrl
    };

    if (editId) {
      setRecords(prev => prev.map(r => r.id === editId ? newRecord : r));
    } else {
      setRecords(prev => [...prev, newRecord]);
    }

    if (odoValue > (activeVehicle?.manualKM || 0)) {
       setVehicles(prev => prev.map(v => v.id === activeVehicle?.id ? { ...v, manualKM: odoValue } : v));
    }
    
    setIsServiceModalOpen(false);
    showToast("Catatan disimpan");
  };

  const openAddServiceModal = () => {
    setEditId('');
    setServiceDate(new Date().toISOString().split('T')[0]);
    setOdometer(activeVehicle ? (activeVehicle.manualKM || '').toString() : '');
    setWorkshop('');
    setServiceItems([{ id: Date.now().toString(), name: '', price: '', type: 'part' }]);
    setAttachmentFile(null);
    setAttachmentUrl('');
    setIsServiceModalOpen(true);
  };

  const openEditRecordModal = (id: string) => {
    const r = records.find(rec => rec.id === id);
    if (!r) return;
    setEditId(r.id);
    setServiceDate(r.date);
    setOdometer(r.odometer.toString());
    setWorkshop(r.workshop || '');
    setAttachmentFile(null);
    setAttachmentUrl(r.attachmentUrl || '');
    if (r.items.length) {
      setServiceItems(r.items.map((i, idx) => ({ id: idx.toString(), name: i.name, price: i.price.toString(), type: i.type || 'part' })));
    } else {
      setServiceItems([{ id: Date.now().toString(), name: '', price: '', type: 'part' }]);
    }
    setIsServiceModalOpen(true);
  };

  const handleDeleteRecord = (id: string) => {
    confirmAction('Hapus Riwayat', 'Anda yakin ingin menghapus riwayat ini?', () => {
      setRecords(prev => prev.filter(r => r.id !== id));
      showToast("Catatan dihapus");
    });
  };

  const handleClearHistory = () => {
    if (!activeVehicle) return;
    confirmAction('Reset Data', 'Anda yakin ingin menghapus semua riwayat motor ini?', () => {
      setRecords(prev => prev.filter(r => r.vehicleId !== activeVehicle.id));
      showToast("Riwayat dibersihkan");
    });
  };

  const downloadPDFReport = () => {
    if (!activeVehicle || vehicleRecords.length === 0) {
      showToast("Tidak ada riwayat untuk diunduh");
      return;
    }
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    
    // Add Title
    doc.setFontSize(18);
    doc.text(`Laporan Riwayat Servis ${activeVehicle.name}`, pageWidth / 2, 22, { align: 'center' });
    
    // Add Subtitle
    doc.setFontSize(11);
    doc.setTextColor(100);
    doc.text(`Diekspor pada ${new Date().toLocaleString('id-ID')}`, pageWidth / 2, 30, { align: 'center' });
    
    // Prepare table data
    const tableBody: any[] = [];
    vehicleRecords.forEach((record, recordIndex) => {
      const parts = record.items;
      const fillColor = recordIndex % 2 === 0 ? [255, 255, 255] : [247, 248, 250];
      
      if (parts.length === 0) {
        tableBody.push([
          { content: (recordIndex + 1).toString(), styles: { fillColor, cellPadding: 3, halign: 'center' } },
          { content: new Date(record.date).toLocaleDateString('id-ID'), styles: { fillColor, cellPadding: 3 } },
          { content: `${formatCurrency(record.odometer)} KM`, styles: { fillColor, cellPadding: 3 } },
          { content: record.workshop || '-', styles: { fillColor, cellPadding: 3 } },
          { content: '-', styles: { fillColor, cellPadding: { left: 2, right: 1, top: 3, bottom: 3 }, halign: 'center' } },
          { content: '-', styles: { fillColor, cellPadding: { left: 1, right: 2, top: 3, bottom: 3 } } },
          { content: `Rp ${formatCurrency(record.totalCost)}`, styles: { fillColor, cellPadding: 3 } }
        ]);
      } else {
        parts.forEach((item, index) => {
          const isFirst = index === 0;
          const isLast = index === parts.length - 1;
          const topPad = isFirst ? 3 : 0.5;
          const bottomPad = isLast ? 3 : 0.5;

          const bulletStyle = { fillColor, cellPadding: { left: 2, right: 1, top: topPad, bottom: bottomPad }, halign: 'right' };
          const textStyle = { fillColor, cellPadding: { left: 1, right: 2, top: topPad, bottom: bottomPad } };

          const rowData = isFirst ? [
            { content: (recordIndex + 1).toString(), rowSpan: parts.length, styles: { fillColor, cellPadding: 3, halign: 'center' } },
            { content: new Date(record.date).toLocaleDateString('id-ID'), rowSpan: parts.length, styles: { fillColor, cellPadding: 3 } },
            { content: `${formatCurrency(record.odometer)} KM`, rowSpan: parts.length, styles: { fillColor, cellPadding: 3 } },
            { content: record.workshop || '-', rowSpan: parts.length, styles: { fillColor, cellPadding: 3 } },
            { content: `\u2022`, styles: bulletStyle as any },
            { content: `${item.name} (Rp ${formatCurrency(item.price)})`, styles: textStyle as any },
            { content: `Rp ${formatCurrency(record.totalCost)}`, rowSpan: parts.length, styles: { fillColor, cellPadding: 3 } }
          ] : [
            { content: `\u2022`, styles: bulletStyle as any },
            { content: `${item.name} (Rp ${formatCurrency(item.price)})`, styles: textStyle as any }
          ];
          tableBody.push(rowData);
        });
      }
    });

    const totalPagesExp = '{total_pages_count_string}';

    autoTable(doc, {
      startY: 40,
      head: [[
        'No',
        'Tanggal', 
        'Odometer', 
        'Bengkel', 
        { content: 'Suku Cadang/Jasa', colSpan: 2 }, 
        'Total Biaya'
      ]],
      body: tableBody,
      styles: { fontSize: 9, cellPadding: 3, valign: 'top' },
      headStyles: { fillColor: [40, 40, 40], halign: 'center' },
      columnStyles: {
        0: { cellWidth: 10, halign: 'center' },
        1: { cellWidth: 25, halign: 'center' },
        2: { cellWidth: 25, halign: 'center' },
        3: { cellWidth: 30, halign: 'center' },
        4: { cellWidth: 5, halign: 'right' },
        5: { cellWidth: 'auto' },
        6: { cellWidth: 30, halign: 'right' }
      },
      didDrawPage: (data) => {
        const pageSize = doc.internal.pageSize;
        const pageWidth = pageSize.width ? pageSize.width : pageSize.getWidth();
        const pageHeight = pageSize.height ? pageSize.height : pageSize.getHeight();
        
        doc.setFontSize(9);
        
        const str1 = "Hal. ";
        const str2 = `${data.pageNumber}`;
        const str3 = " dari ";
        const str4 = totalPagesExp;
        
        doc.setFont("helvetica", "normal");
        const w1 = doc.getTextWidth(str1);
        doc.setFont("helvetica", "bold");
        const w2 = doc.getTextWidth(str2);
        doc.setFont("helvetica", "normal");
        const w3 = doc.getTextWidth(str3);
        doc.setFont("helvetica", "bold");
        const w4 = doc.getTextWidth("99"); 
        
        const totalWidth = w1 + w2 + w3 + w4;
        let startX = (pageWidth - totalWidth) / 2;
        const startY = pageHeight - 10;
        
        doc.setFont("helvetica", "normal");
        doc.text(str1, startX, startY);
        startX += w1;
        
        doc.setFont("helvetica", "bold");
        doc.text(str2, startX, startY);
        startX += w2;
        
        doc.setFont("helvetica", "normal");
        doc.text(str3, startX, startY);
        startX += w3;
        
        doc.setFont("helvetica", "bold");
        doc.text(str4, startX, startY);
      }
    });

    if (typeof doc.putTotalPages === 'function') {
      doc.putTotalPages(totalPagesExp);
    }
    
    const now = new Date();
    const pad = (n: number) => n.toString().padStart(2, '0');
    const yyyy = now.getFullYear();
    const mm = pad(now.getMonth() + 1);
    const dd = pad(now.getDate());
    const hh = pad(now.getHours());
    const min = pad(now.getMinutes());
    const ss = pad(now.getSeconds());
    const dateStr = `${yyyy}${mm}${dd} ${hh}.${min}.${ss}`;
    
    doc.save(`Riwayat Servis ${activeVehicle.name} ${dateStr}.pdf`);
    showToast("Laporan PDF berhasil diunduh");
  };

  const handleAddVehicle = async () => {
    const name = newVehicleName.trim();
    if (!name) return;
    
    setIsVehicleUploading(true);
    let uploadedUrl = '';
    
    if (vehiclePhotoFile) {
        try {
            // Compress and convert to base64 for vehicle icons (avoid Firebase storage issues)
            uploadedUrl = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    const img = new Image();
                    img.onload = () => {
                        const canvas = document.createElement('canvas');
                        const MAX_WIDTH = 150;
                        const MAX_HEIGHT = 150;
                        let width = img.width;
                        let height = img.height;

                        if (width > height) {
                            if (width > MAX_WIDTH) {
                                height *= MAX_WIDTH / width;
                                width = MAX_WIDTH;
                            }
                        } else {
                            if (height > MAX_HEIGHT) {
                                width *= MAX_HEIGHT / height;
                                height = MAX_HEIGHT;
                            }
                        }
                        canvas.width = width;
                        canvas.height = height;
                        const ctx = canvas.getContext('2d');
                        ctx?.drawImage(img, 0, 0, width, height);
                        resolve(canvas.toDataURL('image/jpeg', 0.8));
                    };
                    img.onerror = () => reject(new Error('Image processing failed'));
                    img.src = e.target?.result as string;
                };
                reader.onerror = () => reject(new Error('File reading failed'));
                reader.readAsDataURL(vehiclePhotoFile);
            });
        } catch (err) {
            console.error("Upload error:", err);
            showToast("Gagal memproses foto");
        }
    }

    if (editVehicleId) {
      setVehicles(prev => prev.map(v => {
          if (v.id === editVehicleId) {
              const updated = { ...v, name };
              if (uploadedUrl) {
                  updated.photoUrl = uploadedUrl;
              } else if (!vehiclePhotoPreview) {
                  delete updated.photoUrl;
              }
              return updated;
          }
          return v;
      }));
      setEditVehicleId('');
    } else {
      const usedThemes = vehicles.map(v => v.themeColor);
      const availableThemes = THEMES.filter(t => !usedThemes.includes(t));
      const nextTheme = availableThemes.length > 0 ? availableThemes[Math.floor(Math.random() * availableThemes.length)] : THEMES[Math.floor(Math.random() * THEMES.length)];
      const newVeh = { id: 'veh_' + Date.now(), name, manualKM: 0, themeColor: nextTheme, photoUrl: uploadedUrl };
      setVehicles(prev => [...prev, newVeh]);
      setCurrentVehicleId(newVeh.id);
    }
    setNewVehicleName('');
    setVehiclePhotoFile(null);
    setVehiclePhotoPreview(null);
    if (vehiclePhotoInputRef.current) vehiclePhotoInputRef.current.value = '';
    setIsVehicleUploading(false);
  };

  const handleDeleteVehicle = (id: string) => {
    if (vehicles.length > 1) {
      confirmAction('Hapus Kendaraan', 'Hapus kendaraan ini beserta seluruh riwayatnya?', () => {
        setVehicles(prev => prev.filter(v => v.id !== id));
        setRecords(prev => prev.filter(r => r.vehicleId !== id));
        if (currentVehicleId === id) {
          setCurrentVehicleId(vehicles.find(v => v.id !== id)!.id);
        }
      });
    }
  };

  const triggerFileImport = (accept: string) => {
    if (fileInputRef.current) {
      fileInputRef.current.accept = accept;
      fileInputRef.current.click();
    }
  };

  const handleFileImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    const reader = new FileReader();
    if (file.name.endsWith('.json')) {
      reader.onload = (event) => {
        try {
          const data = JSON.parse(event.target?.result as string);
          if (data.vehicles && data.records) {
            confirmAction('Restore Data', 'Data lama akan ditimpa. Lanjutkan?', () => {
              setVehicles(data.vehicles.map((v: any, idx: number) => ({
                 ...v,
                 themeColor: v.themeColor || THEMES[idx % THEMES.length]
              })));
              setRecords(data.records);
              showToast("Data dipulihkan");
              setIsCloudModalOpen(false);
            });
          }
        } catch { 
           showToast("Format salah", "error"); 
        }
      };
      reader.readAsText(file);
    } else if (file.name.endsWith('.xlsx')) {
      reader.onload = (event) => {
        try {
          const workbook = XLSX.read(new Uint8Array(event.target?.result as ArrayBuffer), { type: 'array' });
          let newVehicles = [...vehicles];
          let newRecords = [...records];
          
          workbook.SheetNames.forEach(name => {
            const data = XLSX.utils.sheet_to_json<any>(workbook.Sheets[name]);
            if (data.length) {
              let veh = newVehicles.find(v => v.name === name);
              if (!veh) {
                const usedThemes = newVehicles.map(v => v.themeColor);
                const availableThemes = THEMES.filter(t => !usedThemes.includes(t));
                const nextTheme = availableThemes.length > 0 ? availableThemes[Math.floor(Math.random() * availableThemes.length)] : THEMES[Math.floor(Math.random() * THEMES.length)];
                veh = { id: 'veh_' + Date.now() + Math.random(), name, manualKM: 0, themeColor: nextTheme };
                newVehicles.push(veh);
              }
              const groups: Record<string, any> = {};
              data.forEach(row => {
                const key = `${row.Tanggal}_${row.Odometer}_${row.Bengkel}`;
                if (!groups[key]) {
                  groups[key] = { 
                    id: 'rec_' + Date.now() + Math.random(), 
                    vehicleId: veh!.id, 
                    date: row.Tanggal, 
                    odometer: row.Odometer, 
                    workshop: row.Bengkel, 
                    items: [], 
                    totalCost: 0 
                  };
                }
                groups[key].items.push({ name: row.Item, price: row.Harga });
                groups[key].totalCost += row.Harga;
              });
              Object.values(groups).forEach(r => {
                if(!newRecords.find(xr => xr.date === r.date && xr.odometer === r.odometer)) {
                  newRecords.push(r);
                }
              });
            }
          });
          setVehicles(newVehicles);
          setRecords(newRecords);
          showToast("Excel diimpor");
          setIsCloudModalOpen(false);
        } catch {
          showToast("Gagal impor", "error");
        }
      };
      reader.readAsArrayBuffer(file);
    }
    e.target.value = '';
  };

  const AccordionItem: React.FC<{ record: ServiceRecord }> = ({ record }) => {
    const [isOpen, setIsOpen] = useState(false);
    
    return (
      <div className={`bg-white rounded-[24px] shadow-sm border border-slate-200 overflow-hidden transition-all duration-500 hover:shadow-md hover:border-slate-300 ${isOpen ? `ring-2 ${theme.ringFocus} shadow-md` : ''}`}>
        <div onClick={() => setIsOpen(!isOpen)} className="p-5 flex items-center justify-between cursor-pointer transition-colors hover:bg-slate-50">
            <div className="flex items-center gap-4 flex-1">
                <div className="min-w-0 flex-1 pl-1">
                    <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">{formatDate(record.date)}</p>
                    
                    <div className="flex flex-wrap gap-x-6 gap-y-2 mt-1.5">
                        <div className="flex flex-col">
                            <span className="text-[8px] font-bold text-slate-400 uppercase leading-none mb-1">Odometer</span>
                            <span className="text-sm font-black text-slate-800">{formatCurrency(record.odometer)} KM</span>
                        </div>
                        <div className="flex flex-col">
                            <span className="text-[8px] font-bold text-slate-400 uppercase leading-none mb-1">Bengkel</span>
                            <span className={`text-sm font-bold ${theme.text} truncate`}>{record.workshop || 'Bengkel Umum'}</span>
                        </div>
                    </div>
                </div>
            </div>

            <div className="flex items-center gap-4 ml-4">
                <div className="text-right">
                    <p className="text-[8px] font-bold text-slate-400 uppercase leading-none mb-1">Total Biaya</p>
                    <p className={`text-sm font-black ${theme.text}`}>Rp {formatCurrency(record.totalCost)}</p>
                </div>
                <div className={`p-2 rounded-full transition-colors ${isOpen ? `${theme.bgLight} ${theme.text}` : 'bg-slate-50 text-slate-400 group-hover:bg-slate-100'}`}>
                    <ChevronDown className={`w-5 h-5 transition-transform duration-300 ${isOpen ? 'rotate-180' : ''}`} strokeWidth={2.5}/>
                </div>
            </div>
        </div>

        <div className={`grid transition-all duration-300 ease-in-out ${isOpen ? 'grid-rows-[1fr] opacity-100' : 'grid-rows-[0fr] opacity-0'}`}>
            <div className="overflow-hidden">
                <div className="px-5 pb-5 pt-1">
                    <div className="pt-4 border-t border-slate-100 space-y-4">
                        <div className="bg-slate-50 rounded-2xl px-5 py-3 border border-slate-100">
                            <p className="text-[9px] font-bold text-slate-500 uppercase tracking-wider mb-3">Rincian Item & Jasa</p>
                            <div className="space-y-1.5">
                                 {record.items.map((i, idx) => {
                                    const highlight = historySearchQuery && i.name.toLowerCase().includes(historySearchQuery.toLowerCase()) ? `${theme.bgLight} ${theme.textLight} rounded px-1 -mx-1` : 'text-slate-700';
                                    return (
                                    <div key={idx} className="flex justify-between items-center py-1.5 text-xs font-semibold border-b border-slate-200/50 last:border-0 hover:bg-slate-100 transition-colors px-2 -mx-2 rounded">
                                        <div className="flex items-center gap-2">
                                            <span className={`text-[8px] px-1.5 py-0.5 rounded uppercase font-bold tracking-widest ${i.type === 'jasa' ? `${theme.bgLight} ${theme.text}` : 'bg-slate-200 text-slate-500'}`}>{i.type || 'PART'}</span>
                                            <span className={highlight}>{i.name}</span>
                                        </div>
                                        <span className="text-slate-500 font-mono text-[10px] pl-4">Rp {formatCurrency(i.price)}</span>
                                    </div>
                                    );
                                })}
                            </div>
                        </div>
                        {record.attachmentUrl && (
                            <div className="bg-slate-50 rounded-2xl px-5 py-3 border border-slate-100">
                                <p className="text-[9px] font-bold text-slate-500 uppercase tracking-wider mb-2">Lampiran Bukti (Klik untuk perbesar)</p>
                                <button type="button" onClick={() => setActiveImagePreview(record.attachmentUrl)} className="block w-full max-h-40 overflow-hidden rounded-xl border border-slate-200 hover:opacity-90 transition-opacity cursor-zoom-in text-left">
                                    <img src={record.attachmentUrl} alt="Bukti Servis" className="w-full object-cover" />
                                </button>
                            </div>
                        )}
                        <div className="flex gap-3 mt-1">
                            <button onClick={() => openEditRecordModal(record.id)} className="flex-1 py-3.5 bg-white border border-slate-200 text-slate-700 rounded-xl text-xs font-bold uppercase tracking-wider hover:bg-slate-50 active:scale-95 transition-all shadow-sm">Ubah Data</button>
                            <button onClick={() => handleDeleteRecord(record.id)} className="flex-1 py-3.5 bg-red-50 border border-red-100 text-red-600 rounded-xl text-xs font-bold uppercase tracking-wider hover:bg-red-100 active:scale-95 transition-all shadow-sm">Hapus</button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen pb-24 bg-slate-50 text-slate-900 relative overflow-x-hidden font-sans">
      <div id="toastContainer" className="fixed bottom-6 left-1/2 -translate-x-1/2 z-[100] flex flex-col gap-2 w-full max-w-sm px-4 text-center pointer-events-none">
        {toasts.map(t => (
          <div key={t.id} className={`${t.type === 'success' ? 'bg-slate-800' : 'bg-red-600'} text-white px-5 py-3.5 rounded-2xl shadow-2xl flex items-center gap-3 toast-enter border border-slate-700 mx-auto pointer-events-auto`}>
            <span className="text-sm font-bold tracking-wide">{t.message}</span>
          </div>
        ))}
      </div>

      {/* Lightbox / Image Preview Modal */}
      <div className={`fixed inset-0 bg-slate-950/85 backdrop-blur-md z-[110] flex items-center justify-center p-4 transition-all duration-300 ${activeImagePreview ? 'opacity-100 visible' : 'opacity-0 invisible'}`}>
          <button onClick={() => setActiveImagePreview(null)} className="absolute top-4 right-4 bg-white/10 hover:bg-white/25 text-white p-3 rounded-full transition-all active:scale-90 z-[120]">
              <X size={24} strokeWidth={2.5} />
          </button>
          <div className="relative max-w-full max-h-[85vh] flex items-center justify-center" onClick={(e) => e.stopPropagation()}>
              {activeImagePreview && (
                  <img src={activeImagePreview} alt="Pratinjau Bukti Servis" className="max-w-full max-h-[85vh] object-contain rounded-2xl shadow-2xl border border-white/10 animate-in zoom-in-95 duration-200" />
              )}
          </div>
          <div className="absolute bottom-6 left-1/2 -translate-x-1/2 text-white/60 text-xs font-bold tracking-wider uppercase pointer-events-none">
              Klik di mana saja untuk menutup
          </div>
          {/* Click background to close */}
          <div className="absolute inset-0 -z-10 cursor-zoom-out" onClick={() => setActiveImagePreview(null)}></div>
      </div>

      <div {...swipeHandlers} className="max-w-md mx-auto px-4 pt-8 md:max-w-4xl relative z-10 w-full overflow-hidden touch-pan-y">
        {/* Header */}
        <header className="flex flex-row items-center justify-between gap-2 mb-8">
            <div className="flex items-center gap-2 sm:gap-3 min-w-0">
                <div className={`p-2 sm:p-2.5 ${theme.bgBase} text-white rounded-[12px] sm:rounded-[14px] shadow-md shrink-0`}>
                    <Wrench className="w-5 h-5 sm:w-6 sm:h-6" strokeWidth={2.5} />
                </div>
                <div className="min-w-0">
                    <h1 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight truncate">MotoJournal</h1>
                    <p className="text-[9px] sm:text-[10px] text-slate-500 font-bold uppercase tracking-widest mt-0.5 truncate">Log Servis Kendaraan</p>
                </div>
            </div>
            <div className="flex gap-1.5 sm:gap-2.5 items-center shrink-0">
                <button onClick={() => setIsReminderModalOpen(true)} title="Sistem Pengingat" className={`bg-white border text-center border-slate-200 text-slate-500 p-2 sm:p-3 rounded-[12px] sm:rounded-[14px] hover:bg-slate-50 hover:${theme.text} active:scale-95 transition-all shadow-sm relative group flex items-center justify-center`}>
                    <Bell className="w-4 h-4 sm:w-5 sm:h-5" strokeWidth={2.5}/>
                    {(componentsNeedingService.length > 0 || componentsApproachingService.length > 0) && (
                      <span className="absolute -top-1 -right-1 w-2.5 h-2.5 bg-red-500 border-2 border-white rounded-full animate-pulse"></span>
                    )}
                </button>
                <button onClick={() => setIsCloudModalOpen(true)} title="Sinkronisasi & Backup" className={`bg-white border text-center border-slate-200 text-slate-500 p-2 sm:p-3 rounded-[12px] sm:rounded-[14px] hover:bg-slate-50 hover:${theme.text} active:scale-95 transition-all shadow-sm relative group flex items-center justify-center`}>
                    <RefreshCw className={`w-4 h-4 sm:w-5 sm:h-5 ${isSyncing ? `animate-spin ${theme.text}` : ''}`} strokeWidth={2.5}/>
                </button>
                <button onClick={() => setIsVehicleModalOpen(true)} title="Daftar Kendaraan" className={`bg-white border border-slate-200 text-slate-500 p-2 sm:p-3 rounded-[12px] sm:rounded-[14px] hover:bg-slate-50 hover:${theme.text} active:scale-95 transition-all shadow-sm flex items-center justify-center`}>
                    <Grid className="w-4 h-4 sm:w-5 sm:h-5" strokeWidth={2.5} />
                </button>
                <button onClick={openAddServiceModal} title="Entri Servis Baru" className={`${theme.bgBase} text-white p-2 sm:p-3 rounded-[12px] sm:rounded-[14px] shadow-lg ${theme.shadowLg} ${theme.bgHover} hover:shadow-xl hover:-translate-y-0.5 active:scale-95 transition-all flex items-center justify-center`}>
                    <Plus className="w-4 h-4 sm:w-5 sm:h-5" strokeWidth={3} />
                </button>
            </div>
        </header>

        {/* Vehicle Tabs */}
        <div className="flex gap-2 overflow-x-auto pb-4 mb-4 custom-scroll no-scrollbar -mx-4 px-4 md:mx-0 md:px-0">
          {vehicles.map(v => (
            <button key={v.id} onClick={() => setCurrentVehicleId(v.id)} className={`px-4 py-2 rounded-full text-sm font-bold whitespace-nowrap transition-all duration-300 active:scale-95 flex items-center gap-2 ${v.id === currentVehicleId ? `${getThemeClasses(v.themeColor).bgBase} text-white shadow-md ring-2 ring-slate-800/20 ring-offset-2 ring-offset-slate-50` : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-100 hover:text-slate-900 shadow-sm'}`}>
              {v.photoUrl && <img src={v.photoUrl} alt={v.name} className="w-6 h-6 rounded-full object-cover border border-white/20 shadow-sm" />}
              {v.name}
            </button>
          ))}
        </div>

        {/* Main Tabs */}
        <div className="bg-slate-200/50 p-1.5 rounded-3xl flex gap-1 mb-6 shadow-inner border border-slate-200 relative">
            <button onClick={() => handleTabChange('home')} className={`flex-1 py-2.5 px-3 rounded-2xl text-[11px] font-black uppercase tracking-wider transition-all duration-300 ${currentMainTab === 'home' ? `bg-white ${theme.textLight} shadow-sm border border-slate-100` : 'text-slate-500 hover:text-slate-700'}`}>Beranda</button>
            <button onClick={() => handleTabChange('analysis')} className={`flex-1 py-2.5 px-3 rounded-2xl text-[11px] font-black uppercase tracking-wider transition-all duration-300 ${currentMainTab === 'analysis' ? `bg-white ${theme.textLight} shadow-sm border border-slate-100` : 'text-slate-500 hover:text-slate-700'}`}>Analisis</button>
            <button onClick={() => handleTabChange('history')} className={`flex-1 py-2.5 px-3 rounded-2xl text-[11px] font-black uppercase tracking-wider transition-all duration-300 ${currentMainTab === 'history' ? `bg-white ${theme.textLight} shadow-sm border border-slate-100` : 'text-slate-500 hover:text-slate-700'}`}>Riwayat</button>
        </div>

        {/* Home Section */}
        <div className="relative overflow-hidden w-full min-h-[500px]">
          <AnimatePresence mode="wait" initial={false}>
            {currentMainTab === 'home' && (
              <motion.section
                key="home"
                initial={{ opacity: 0, x: tabDirection > 0 ? 30 : -30 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: tabDirection > 0 ? -30 : 30 }}
                transition={{ duration: 0.2, ease: "easeOut" }}
                className="space-y-6"
              >
              {componentsNeedingService.length > 0 && (
                  <div className="bg-red-50 border border-red-200 p-4 rounded-[20px] flex items-start gap-4 shadow-sm animate-in fade-in zoom-in-95 duration-300">
                      <div className="bg-red-100 p-2.5 rounded-xl text-red-600 shadow-sm border border-red-200">
                          <AlertTriangle size={20} strokeWidth={2.5} />
                      </div>
                      <div className="flex-1 mt-0.5">
                          <h4 className="text-sm font-black text-red-800 tracking-tight mb-1">Perhatian: Waktunya Servis!</h4>
                          <p className="text-xs font-medium text-red-700/90 leading-relaxed mb-3">
                              {componentsNeedingService.length} komponen telah mencapai batas interval dan perlu segera diperiksa: <span className="font-bold">{componentsNeedingService.map(c => c.name).join(', ')}</span>.
                          </p>
                          <button onClick={() => handleTabChange('analysis')} className="text-xs font-bold bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-xl shadow-sm transition-all active:scale-95">Lihat Detail Analisis</button>
                      </div>
                  </div>
              )}
              {componentsApproachingService.length > 0 && (
                  <div className="bg-amber-50 border border-amber-200 p-4 rounded-[20px] flex items-start gap-4 shadow-sm animate-in fade-in zoom-in-95 duration-300">
                      <div className="bg-amber-100 p-2.5 rounded-xl text-amber-600 shadow-sm border border-amber-200">
                          <Bell size={20} strokeWidth={2.5} />
                      </div>
                      <div className="flex-1 mt-0.5">
                          <h4 className="text-sm font-black text-amber-800 tracking-tight mb-1">Pemberitahuan: Mendekati Batas Servis!</h4>
                          <p className="text-xs font-medium text-amber-700/90 leading-relaxed mb-3">
                              {componentsApproachingService.length} komponen hampir mencapai batas interval penggantian (dalam {formatCurrency(reminderThreshold)} KM): <span className="font-bold">{componentsApproachingService.map(c => c.name).join(', ')}</span>.
                          </p>
                          <button onClick={() => handleTabChange('analysis')} className="text-xs font-bold bg-amber-600 hover:bg-amber-700 text-white px-4 py-2 rounded-xl shadow-sm transition-all active:scale-95">Lihat Detail Analisis</button>
                      </div>
                  </div>
              )}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 border-b border-slate-200/60 pb-6 mb-6">
                  <div className="bg-white p-6 rounded-[24px] shadow-sm border border-slate-200 relative group transition-all duration-300 hover:shadow-md hover:border-slate-300">
                      <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1 shadow-none transition-colors">Odometer Saat Ini</p>
                      <h3 className="text-3xl font-black text-slate-800 tracking-tight transition-colors">{formatCurrency(activeVehicle?.manualKM || 0)} <span className="text-sm font-bold text-slate-500 tracking-normal inline-block ml-1">KM</span></h3>
                      <button onClick={() => { setManualKMInput(activeVehicle?.manualKM?.toString() || '0'); setIsKMModalOpen(true); }} className={`absolute top-4 right-4 p-2.5 bg-slate-50 text-slate-400 rounded-xl lg:opacity-0 group-hover:opacity-100 transition-all duration-300 hover:bg-slate-100 hover:${theme.text} active:scale-95 border border-slate-200`}>
                          <Edit2 size={16} strokeWidth={2.5} />
                      </button>
                  </div>
                  <div className="bg-white p-6 rounded-[24px] shadow-sm border border-slate-200 transition-all duration-300 hover:shadow-md hover:border-slate-300">
                      <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1 shadow-none">Servis Terakhir</p>
                      <h3 className="text-2xl font-black text-slate-800 tracking-tight">{formatCurrency(latestRecordByOdo ? latestRecordByOdo.odometer : 0)} <span className="text-sm font-bold text-slate-500 tracking-normal inline-block ml-1">KM</span></h3>
                  </div>
                  <div className="bg-white p-6 rounded-[24px] shadow-sm border border-slate-200 transition-all duration-300 hover:shadow-md hover:border-slate-300">
                      <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1 shadow-none">Tanggal Terakhir</p>
                      <h3 className="text-lg font-bold text-slate-800 mt-2">{latestRecordByOdo ? formatDate(latestRecordByOdo.date) : "-"}</h3>
                  </div>
              </div>

              <div className={`${theme.bgBase} border border-white/10 p-8 rounded-[32px] text-white shadow-xl ${theme.shadowHover} relative overflow-hidden group hover:shadow-2xl transition-all duration-500`}>
                  <div className="absolute top-[-30%] right-[-10%] w-[60%] h-[150%] bg-white/10 blur-3xl rounded-full mix-blend-overlay pointer-events-none"></div>
                  <div className="absolute bottom-[-50%] left-[-20%] w-[50%] h-[100%] bg-white/10 blur-2xl rounded-full pointer-events-none"></div>
                  <div className="relative z-10">
                      <h2 className="text-3xl font-black mb-2 tracking-tight">Halo, <span className="opacity-90">{activeVehicle?.name}</span>.</h2>
                      <p className="text-sm font-medium text-white/90 leading-relaxed mb-8 max-w-sm">Pantau terus performa kendaraan Anda agar selalu prima di jalanan.</p>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                          <button onClick={openAddServiceModal} className="bg-white/10 hover:bg-white/20 active:scale-95 border border-white/20 backdrop-blur-md p-5 rounded-[24px] transition-all duration-300 text-left group/btn shadow-sm">
                              <p className="text-[10px] font-bold uppercase tracking-widest opacity-80 mb-1.5 text-white/80">Aksi Cepat</p>
                              <p className="text-sm font-bold tracking-wide">Catat Servis</p>
                          </button>
                          <button onClick={() => handleTabChange('analysis')} className="bg-white/10 hover:bg-white/20 active:scale-95 border border-white/20 backdrop-blur-md p-5 rounded-[24px] transition-all duration-300 text-left group/btn shadow-sm">
                              <p className="text-[10px] font-bold uppercase tracking-widest opacity-80 mb-1.5 text-white/80">Cek Mesin</p>
                              <p className="text-sm font-bold tracking-wide">Lihat Analisa</p>
                          </button>
                      </div>
                  </div>
              </div>
              </motion.section>
            )}

            {/* Analysis Section */}
            {currentMainTab === 'analysis' && (
              <motion.section
                key="analysis"
                initial={{ opacity: 0, x: tabDirection > 0 ? 30 : -30 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: tabDirection > 0 ? -30 : 30 }}
                transition={{ duration: 0.2, ease: "easeOut" }}
                className="space-y-5"
              >
              {expenseData.length > 0 && (
                <div className="bg-white border border-slate-200 p-5 rounded-[24px] shadow-sm mb-6">
                  <h2 className="text-[11px] font-black text-slate-500 uppercase tracking-widest mb-4">Grafik Pengeluaran</h2>
                  <div className="h-44 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={expenseData} margin={{ top: 0, right: 0, left: 0, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} dy={10} />
                        <YAxis hide />
                        <Tooltip 
                          cursor={{ fill: '#f1f5f9' }}
                          content={({ active, payload }) => {
                            if (active && payload && payload.length) {
                              return (
                                <div className="bg-slate-800 text-white text-xs font-bold px-3 py-2 rounded-xl shadow-lg">
                                  Rp {formatCurrency(payload[0].value as number)}
                                </div>
                              );
                            }
                            return null;
                          }}
                        />
                        <Bar dataKey="cost" fill="#4f46e5" radius={[4, 4, 0, 0]} maxBarSize={40} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              )}

              {odometerData.length > 1 && (
                <div className="bg-white border border-slate-200 p-5 rounded-[24px] shadow-sm mb-6">
                  <h2 className="text-[11px] font-black text-slate-500 uppercase tracking-widest mb-4">Tren Odometer</h2>
                  <div className="h-44 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={odometerData} margin={{ top: 10, right: 10, left: 10, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="date" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} dy={10} />
                        <YAxis domain={['auto', 'auto']} hide />
                        <Tooltip 
                          cursor={{ stroke: '#94a3b8', strokeWidth: 1, strokeDasharray: '3 3' }}
                          content={({ active, payload }) => {
                            if (active && payload && payload.length) {
                              return (
                                <div className="bg-slate-800 text-white text-xs font-bold px-3 py-2 rounded-xl shadow-lg">
                                  {formatCurrency(payload[0].value as number)} KM
                                </div>
                              );
                            }
                            return null;
                          }}
                        />
                        <Line type="monotone" dataKey="km" stroke="#0ea5e9" strokeWidth={3} dot={{ r: 4, fill: '#0ea5e9', strokeWidth: 2, stroke: '#fff' }} activeDot={{ r: 6, strokeWidth: 0 }} />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              )}

              <div className="flex items-center justify-between px-1 border-b border-slate-200/60 pb-3 mb-2">
                  <h2 className="text-[11px] font-black text-slate-500 uppercase tracking-widest">Kesehatan Komponen</h2>
                  <div className="flex gap-4">
                      <div className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-emerald-500 shadow-sm shadow-emerald-500/50"></span><span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Aman</span></div>
                      <div className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-red-500 shadow-sm shadow-red-500/50 animate-pulse"></span><span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Perlu Servis</span></div>
                  </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {INTEL_CONFIG.map(cfg => {
                    const matches = vehicleRecords.filter(r => 
                      r.items.some(item => cfg.kws.some(kw => item.name.toLowerCase().includes(kw)))
                    );
                    const lastRec = matches.length ? matches[0] : null;
                    const lastKM = lastRec ? lastRec.odometer : 0;
                    const currentKM = activeVehicle?.manualKM || 0;
                    const nextKM = lastKM + cfg.interval;
                    const remKM = nextKM - currentKM;
                    let progress = 0;
                    if (lastRec) {
                        const usedKM = currentKM - lastKM;
                        progress = Math.min(Math.max((usedKM / cfg.interval) * 100, 0), 100);
                    }
                    const isUrgent = remKM <= 0;
                    
                    const healthPercent = Math.round(100 - progress);
                    const lastServiceInfo = lastRec 
                        ? `Terakhir: ${formatCurrency(lastRec.odometer)} KM (${formatDate(lastRec.date)})`
                        : 'Belum ada catatan';

                    const daysRem = Math.round(Math.max(0, remKM) / dailyAverageKM);
                    const estDate = new Date();
                    estDate.setDate(estDate.getDate() + daysRem);
                    
                    return (
                      <div key={cfg.id} className="bg-white border border-slate-200 p-6 rounded-[24px] shadow-sm transition-all duration-300 hover:shadow-md hover:border-slate-300">
                          <div className="flex justify-between items-start">
                              <div>
                                  <p className={`text-[11px] font-black text-slate-400 uppercase tracking-widest mb-1.5`}>{cfg.name}</p>
                                  <h3 className="text-2xl font-black text-slate-800 tracking-tight">
                                      {isUrgent ? <span className="text-red-500">Servis Segera!</span> : <>{formatCurrency(remKM)} <span className="text-sm font-bold text-slate-500 tracking-normal ml-1 inline-block">KM Lagi</span></>}
                                  </h3>
                                  <p className="text-[10px] font-bold text-slate-500 mt-1">{lastServiceInfo}</p>
                              </div>
                              <div className={`p-3 bg-slate-50 shadow-inner rounded-xl border border-slate-200 text-slate-400`}>
                                   <Wrench size={20} strokeWidth={2.5} />
                              </div>
                          </div>
                          
                          <div className="mt-4 pt-4 border-t border-slate-100 flex justify-between items-center bg-slate-50/50 p-3 rounded-xl">
                              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Estimasi Ganti</span>
                              <span className={`text-xs font-black ${theme.text}`}>{isUrgent ? 'Sekarang' : formatDate(estDate.toISOString().split('T')[0])}</span>
                          </div>

                          <div className="space-y-2.5 mt-5">
                              <div className="flex justify-between text-[10px] font-bold uppercase tracking-widest text-slate-500 group">
                                  <span>Kesehatan Part</span>
                                  <span className={healthPercent <= 20 ? 'text-red-500' : 'text-slate-700'}>{healthPercent}%</span>
                              </div>
                              <div className="h-2.5 w-full bg-slate-100 rounded-full overflow-hidden border border-slate-200/50 shadow-inner">
                                  <div className={`h-full transition-all duration-1000 ease-out ${isUrgent ? 'bg-red-500' : (remKM <= 500 ? 'bg-orange-500' : `bg-emerald-500`)}`} style={{ width: `${healthPercent}%` }}></div>
                              </div>
                          </div>
                      </div>
                    );
                  })}
              </div>
              </motion.section>
            )}

            {/* History Section */}
            {currentMainTab === 'history' && (
              <motion.section
                key="history"
                initial={{ opacity: 0, x: tabDirection > 0 ? 30 : -30 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: tabDirection > 0 ? -30 : 30 }}
                transition={{ duration: 0.2, ease: "easeOut" }}
                className="space-y-5"
              >
              <div className="flex flex-col gap-5 mb-2">
                  <div className="flex flex-row items-center justify-between px-1 gap-2">
                      <h2 className="font-black text-slate-900 text-lg sm:text-2xl tracking-tight leading-tight truncate">Catatan <span className={`${theme.text}`}>{activeVehicle?.name}</span></h2>
                      <div className="flex items-center gap-3 sm:gap-4 shrink-0">
                          <button onClick={downloadPDFReport} className="text-[9px] sm:text-[10px] font-bold uppercase tracking-widest text-slate-500 hover:text-slate-700 hover:underline hover:underline-offset-2 active:scale-95 transition-all">Unduh PDF</button>
                          <button onClick={handleClearHistory} className="text-[9px] sm:text-[10px] font-bold uppercase tracking-widest text-red-500 hover:text-red-600 hover:underline hover:underline-offset-2 active:scale-95 transition-all">Kosongkan</button>
                      </div>
                  </div>
                  
                  <div className="flex flex-col gap-3">
                      <div className="relative group">
                          <div className={`absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none text-slate-400 group-focus-within:${theme.text} transition-colors`}>
                              <Search size={18} strokeWidth={2.5}/>
                          </div>
                          <input type="text" value={historySearchQuery} onChange={e => setHistorySearchQuery(e.target.value)} placeholder="Cari suku cadang, bengkel..." className={`w-full bg-white border border-slate-200 text-slate-800 placeholder:text-slate-400 py-3.5 pl-12 pr-4 rounded-[16px] text-sm font-semibold outline-none ${theme.borderFocus} ${theme.ringFocus} transition-all duration-300 shadow-sm hover:border-slate-300`} />
                      </div>
                      <div className="grid grid-cols-2 gap-3">
                          <select value={historyMonthFilter} onChange={e => setHistoryMonthFilter(e.target.value)} className={`bg-white border border-slate-200 text-slate-700 text-xs font-bold py-3.5 pl-4 pr-8 rounded-[16px] outline-none ${theme.ringFocus} transition-all shadow-sm appearance-none cursor-pointer hover:border-slate-300`}>
                              <option value="">Semua Waktu</option>
                              {Array.from(new Set(vehicleRecords.map(r => r.date.substring(0, 7)))).sort().reverse().map(ym => (
                                  <option key={ym} value={ym}>{new Date(`${ym}-01`).toLocaleString('default', { month: 'long', year: 'numeric' })}</option>
                              ))}
                          </select>
                          <select value={historyPriceFilter} onChange={e => setHistoryPriceFilter(e.target.value)} className={`bg-white border border-slate-200 text-slate-700 text-xs font-bold py-3.5 pl-4 pr-8 rounded-[16px] outline-none ${theme.ringFocus} transition-all shadow-sm appearance-none cursor-pointer hover:border-slate-300`}>
                              <option value="">Semua Harga</option>
                              <option value="0-50000">&lt; Rp 50.000</option>
                              <option value="50000-250000">Rp 50rb - 250rb</option>
                              <option value="250000-1000000">Rp 250rb - 1jt</option>
                              <option value="1000000-999999999">&gt; Rp 1jt</option>
                          </select>
                      </div>
                  </div>
              </div>

              <div className="flex flex-col gap-4">
                {vehicleRecords.length === 0 ? (
                  <div className="py-24 text-center text-slate-500">
                      <div className="w-16 h-16 bg-slate-100 rounded-full flex items-center justify-center mx-auto mb-4 text-slate-400">
                           <Grid size={24} strokeWidth={2.5}/>
                      </div>
                      <p className="text-sm font-semibold text-slate-600">Belum ada riwayat perawatan.</p>
                      <p className="text-xs text-slate-400 mt-1">Catatan servis akan tampil di sini.</p>
                  </div>
                ) : filteredRecords.length === 0 ? (
                  <div className="py-20 text-center text-slate-500">
                      <p className="text-sm font-semibold text-slate-600">Tidak ada catatan yang sesuai pencarian.</p>
                  </div>
                ) : (
                  filteredRecords.map(record => (
                     <AccordionItem key={record.id} record={record} />
                  ))
                )}
              </div>
              </motion.section>
            )}
          </AnimatePresence>
        </div>

      </div>

      {/* KM Modal */}
      <div className={`fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-[60] flex items-end md:items-center justify-center transition-all duration-300 ${isKMModalOpen ? 'opacity-100 visible' : 'opacity-0 invisible'}`}>
        <div className={`bg-white rounded-t-3xl md:rounded-3xl shadow-2xl border border-slate-100 w-full max-w-sm overflow-hidden transform transition-all duration-300 ${isKMModalOpen ? 'translate-y-0 scale-100' : 'translate-y-12 md:translate-y-0 md:scale-95'}`}>
            <div className="p-5 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                <h3 className="text-lg font-black text-slate-800 tracking-tight">Update Odometer</h3>
                <button onClick={() => setIsKMModalOpen(false)} className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-200/50 rounded-xl transition-all"><X size={20} strokeWidth={2.5}/></button>
            </div>
            <div className="p-6 space-y-5">
                <div className="space-y-1 group">
                    <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest px-1">Odometer Saat Ini (KM)</label>
                    <input type="number" value={manualKMInput} onChange={e => setManualKMInput(e.target.value)} placeholder="0" className={`w-full px-4 py-4 bg-white border border-slate-200 ${theme.borderFocus} shadow-sm rounded-[16px] text-xl font-black text-slate-800 placeholder:text-slate-300 outline-none ${theme.ringFocus} transition-all duration-300`} />
                </div>
                <button onClick={saveManualKM} className={`w-full ${theme.bgBase} ${theme.bgHover} text-white font-bold py-4 rounded-[16px] shadow-lg ${theme.shadowHover} active:scale-95 transition-all duration-300`}>Simpan Odometer</button>
            </div>
        </div>
      </div>

      {/* Service Modal */}
      <div className={`fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-end md:items-center justify-center transition-all duration-300 ${isServiceModalOpen ? 'opacity-100 visible' : 'opacity-0 invisible'}`}>
        <div className={`bg-white rounded-t-3xl md:rounded-3xl shadow-2xl border border-slate-100 w-full max-w-xl overflow-hidden transform transition-all duration-300 ${isServiceModalOpen ? 'translate-y-0 scale-100' : 'translate-y-12 md:translate-y-0 md:scale-95'}`}>
            <div className="p-5 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                <h3 className="text-lg font-black text-slate-800 tracking-tight">{editId ? 'Perbarui Catatan' : 'Entri Servis Baru'}</h3>
                <button onClick={() => setIsServiceModalOpen(false)} className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-200/50 rounded-xl transition-all">
                    <X size={20} strokeWidth={2.5}/>
                </button>
            </div>
            <form onSubmit={handleServiceSubmit} className="p-4 sm:p-6 space-y-4 sm:space-y-5 max-h-[85vh] overflow-y-auto custom-scroll pb-6">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div className="space-y-1 w-full">
                        <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest px-1">Tanggal</label>
                        <input type="date" required value={serviceDate} onChange={e => setServiceDate(e.target.value)} className={`w-full px-4 py-3.5 bg-white border border-slate-200 ${theme.borderFocus} shadow-sm rounded-xl text-sm font-medium ${theme.ringFocus} outline-none text-slate-800 transition-all duration-300`} />
                    </div>
                    <div className="space-y-1 w-full">
                        <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest px-1">Odometer Servis</label>
                        <input type="number" required value={odometer} onChange={e => setOdometer(e.target.value)} placeholder="0" className={`w-full px-4 py-3.5 bg-white border border-slate-200 ${theme.borderFocus} shadow-sm rounded-xl text-sm font-medium ${theme.ringFocus} outline-none text-slate-800 transition-all duration-300`} />
                    </div>
                </div>
                <div className="space-y-1">
                    <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest px-1">Bengkel</label>
                    <input type="text" value={workshop} onChange={e => setWorkshop(e.target.value)} placeholder="Nama Bengkel" className={`w-full px-4 py-3.5 bg-white border border-slate-200 ${theme.borderFocus} shadow-sm rounded-xl text-sm font-medium ${theme.ringFocus} outline-none text-slate-800 transition-all duration-300`} />
                </div>
                <div className="space-y-3">
                    <div className="flex justify-between items-center px-1">
                        <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Rincian Komponen / Jasa</label>
                        <button type="button" onClick={handleAddServiceItem} className={`text-xs ${theme.bgLight} ${theme.textLight} px-4 py-2 rounded-xl font-bold hover:brightness-95 transition-all border ${theme.border}`}>+ Item</button>
                    </div>
                    <div className="space-y-3 max-h-60 overflow-y-auto custom-scroll p-1 -m-1">
                      {serviceItems.map((item, idx) => (
                         <div key={item.id} className={`flex flex-col gap-3 p-4 bg-slate-50 border border-slate-200 rounded-2xl relative transition-all duration-300 group-focus-within:border-slate-300 focus-within:bg-white focus-within:shadow-sm ${theme.borderFocus}`}>
                             <input type="text" placeholder="Nama part/jasa" value={item.name} required onChange={e => {
                               const newItems = [...serviceItems]; newItems[idx].name = e.target.value; setServiceItems(newItems);
                             }} className={`w-full pr-8 bg-transparent border-b border-slate-200 ${theme.borderFocus} outline-none text-sm font-semibold py-1.5 text-slate-800 placeholder:text-slate-400 transition-colors`}/>
                             <div className="flex items-center justify-between mt-1">
                                 <div className="flex bg-slate-200/50 p-1 rounded-xl">
                                    <button type="button" onClick={() => { const newItems = [...serviceItems]; newItems[idx].type = 'part'; setServiceItems(newItems); }} className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-all ${item.type !== 'jasa' ? `bg-white ${theme.text} shadow-sm` : 'text-slate-500 hover:text-slate-700'}`}>Part</button>
                                    <button type="button" onClick={() => { const newItems = [...serviceItems]; newItems[idx].type = 'jasa'; setServiceItems(newItems); }} className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-all ${item.type === 'jasa' ? `bg-white ${theme.text} shadow-sm` : 'text-slate-500 hover:text-slate-700'}`}>Jasa</button>
                                 </div>
                                 <input type="number" placeholder="Rp 0" value={item.price} onChange={e => {
                                   const newItems = [...serviceItems]; newItems[idx].price = e.target.value; setServiceItems(newItems);
                                 }} className={`bg-transparent text-right outline-none text-sm font-mono font-bold w-1/2 text-slate-800 placeholder:text-slate-300 transition-colors focus:${theme.text} py-1`}/>
                             </div>
                             {serviceItems.length > 1 && (
                                <button type="button" onClick={() => setServiceItems(prev => prev.filter(i => i.id !== item.id))} className="absolute right-3 top-3 bg-white/80 backdrop-blur-sm shadow-md p-2 rounded-xl text-slate-400 hover:text-red-500 hover:bg-white border border-slate-200 transition-all active:scale-90">
                                  <X size={16} strokeWidth={2.5} />
                                </button>
                             )}
                         </div>
                      ))}
                    </div>
                </div>
                
                <div className="space-y-2">
                    <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest px-1">Lampiran Foto BUKTI (Opsional)</label>
                    <div className="flex items-center gap-3">
                        <label className="flex items-center gap-2 px-4 py-3 bg-white border border-slate-200 text-slate-600 rounded-xl cursor-pointer hover:bg-slate-50 transition-all shadow-sm">
                            <Paperclip size={16} />
                            <span className="text-xs font-bold truncate max-w-[150px] sm:max-w-[200px]">{attachmentFile ? attachmentFile.name : (attachmentUrl ? 'Ganti Foto' : 'Pilih Foto')}</span>
                            <input type="file" accept="image/*" className="hidden" onChange={e => { if (e.target.files && e.target.files[0]) setAttachmentFile(e.target.files[0]); }} />
                        </label>
                        {(attachmentFile || attachmentUrl) && (
                            <div className="h-10 w-10 relative overflow-hidden rounded-lg border border-slate-200">
                                <img src={attachmentFile ? URL.createObjectURL(attachmentFile) : attachmentUrl} alt="lampiran" className="object-cover w-full h-full" />
                            </div>
                        )}
                        {(attachmentFile || attachmentUrl) && (
                            <button type="button" onClick={() => { setAttachmentFile(null); setAttachmentUrl(''); }} className="text-red-500 p-2 hover:bg-red-50 rounded-lg transition-colors">
                                <X size={16} strokeWidth={2.5}/>
                            </button>
                        )}
                    </div>
                </div>

                <div className="pt-5 border-t border-slate-100 mt-2">
                    <div className="flex justify-between items-center mb-6 px-1">
                        <span className="text-xs font-bold text-slate-500 uppercase tracking-widest">Total Biaya</span>
                        <span className={`text-2xl font-black ${theme.text} tracking-tight`}>Rp {formatCurrency(calculateFormTotal())}</span>
                    </div>
                    <button type="submit" disabled={isUploading} className={`w-full ${theme.bgBase} disabled:opacity-70 flex justify-center items-center gap-2 ${theme.bgHover} text-white font-bold py-4 rounded-[16px] shadow-lg ${theme.shadowHover} active:scale-95 transition-all duration-300`}>
                        {isUploading ? <><RefreshCw size={18} className="animate-spin" /> Mengupload...</> : 'Simpan Catatan'}
                    </button>
                </div>
            </form>
        </div>
      </div>

      {/* Vehicles Modal */}
      <div className={`fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-end md:items-center justify-center transition-all duration-300 ${isVehicleModalOpen ? 'opacity-100 visible' : 'opacity-0 invisible'}`}>
        <div className={`bg-white rounded-t-3xl md:rounded-3xl shadow-2xl border border-slate-100 w-full max-w-xl overflow-hidden transform transition-all duration-300 ${isVehicleModalOpen ? 'translate-y-0 scale-100' : 'translate-y-12 md:translate-y-0 md:scale-95'}`}>
            <div className="p-5 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                <h3 className="text-lg font-black text-slate-800 tracking-tight">Daftar Kendaraan</h3>
                <button onClick={() => setIsVehicleModalOpen(false)} className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-200/50 rounded-xl transition-all"><X size={20} strokeWidth={2.5}/></button>
            </div>
            <div className="p-6 space-y-6">
                <div className="space-y-2">
                    <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest px-1">Tambah / Edit Kendaraan</label>
                    <div className="flex flex-col gap-3">
                        <div className="flex gap-2">
                            <input type="text" value={newVehicleName} onChange={e => setNewVehicleName(e.target.value)} placeholder="Nama Kendaraan" className={`flex-1 px-4 py-3 bg-white border border-slate-200 ${theme.borderFocus} shadow-sm rounded-xl text-sm font-medium text-slate-800 placeholder:text-slate-400 ${theme.ringFocus} outline-none transition-all duration-300`} />
                            <button onClick={handleAddVehicle} disabled={isVehicleUploading} className="bg-slate-800 hover:bg-slate-900 disabled:opacity-70 text-white px-6 rounded-xl font-bold text-sm shadow-md active:scale-95 transition-all duration-300">
                                {isVehicleUploading ? <RefreshCw className="w-5 h-5 animate-spin" /> : editVehicleId ? 'Update' : 'Simpan'}
                            </button>
                        </div>
                        <div className="flex items-center gap-3">
                            <button type="button" onClick={() => vehiclePhotoInputRef.current?.click()} className={`flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 rounded-xl text-sm font-medium text-slate-600 hover:bg-slate-50 transition-all ${theme.borderFocus}`}>
                                <Camera size={18} />
                                {vehiclePhotoFile ? 'Ganti Foto' : 'Pilih Foto (Opsional)'}
                            </button>
                            {vehiclePhotoPreview && (
                                <div className="relative w-10 h-10 rounded-lg overflow-hidden border border-slate-200 shadow-sm">
                                    <img src={vehiclePhotoPreview} alt="Preview" className="w-full h-full object-cover" />
                                    <button type="button" onClick={() => { setVehiclePhotoFile(null); setVehiclePhotoPreview(null); if(vehiclePhotoInputRef.current) vehiclePhotoInputRef.current.value = ''; }} className="absolute -top-1 -right-1 bg-red-500 text-white rounded-full p-0.5"><X size={12}/></button>
                                </div>
                            )}
                            <input type="file" accept="image/*" ref={vehiclePhotoInputRef} className="hidden" onChange={(e) => {
                                const file = e.target.files?.[0];
                                if (file) {
                                    setVehiclePhotoFile(file);
                                    setVehiclePhotoPreview(URL.createObjectURL(file));
                                }
                            }} />
                        </div>
                    </div>
                </div>
                <div className="space-y-2 max-h-60 overflow-y-auto custom-scroll p-1 -m-1">
                  {vehicles.map(v => (
                     <div key={v.id} className="flex items-center justify-between p-4 bg-white border border-slate-200 hover:border-slate-300 hover:shadow-sm rounded-2xl transition-all duration-300">
                        <div className="flex items-center gap-3">
                            {v.photoUrl ? (
                                <img src={v.photoUrl} alt={v.name} className="w-10 h-10 rounded-full object-cover border border-slate-200" />
                            ) : (
                                <div className={`w-10 h-10 rounded-full flex items-center justify-center text-white font-bold ${getThemeClasses(v.themeColor).bgBase}`}>
                                    {v.name.substring(0, 2).toUpperCase()}
                                </div>
                            )}
                            <span className="text-sm font-bold text-slate-800">{v.name}</span>
                        </div>
                        <div className="flex gap-1.5">
                            <button onClick={() => { setEditVehicleId(v.id); setNewVehicleName(v.name); setVehiclePhotoPreview(v.photoUrl || null); setVehiclePhotoFile(null); }} className={`text-slate-500 p-2 hover:${theme.bgLight} hover:${theme.text} border border-transparent hover:${theme.border} rounded-xl transition-all`}><Edit2 size={16} strokeWidth={2.5} /></button>
                            <button onClick={() => handleDeleteVehicle(v.id)} className="text-slate-500 p-2 hover:bg-red-50 hover:text-red-600 border border-transparent hover:border-red-100 rounded-xl transition-all"><Trash2 size={16} strokeWidth={2.5} /></button>
                        </div>
                    </div>
                  ))}
                </div>
            </div>
        </div>
      </div>

      {/* Sync Modal */}
      <div className={`fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-end md:items-center justify-center transition-all duration-300 ${isCloudModalOpen ? 'opacity-100 visible' : 'opacity-0 invisible'}`}>
        <div className={`bg-white rounded-t-3xl md:rounded-3xl shadow-2xl border border-slate-100 w-full max-w-sm overflow-hidden transform transition-all duration-300 ${isCloudModalOpen ? 'translate-y-0 scale-100' : 'translate-y-12 md:translate-y-0 md:scale-95'}`}>
            <div className="p-5 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                <h3 className="text-lg font-black text-slate-800 tracking-tight">Ekspor & Backup</h3>
                <button onClick={() => setIsCloudModalOpen(false)} className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-200/50 rounded-xl transition-all"><X size={20} strokeWidth={2.5}/></button>
            </div>
            <div className="p-6 space-y-4">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <button onClick={() => exportToExcel(records, vehicles, showToast)} className="bg-emerald-50 border border-emerald-100 text-emerald-700 py-3 rounded-[14px] font-bold text-[11px] uppercase tracking-wider flex items-center justify-center gap-2 hover:bg-emerald-100 active:scale-95 transition-all">Ekspor Excel</button>
                    <button onClick={() => triggerFileImport('.xlsx')} className="bg-white border border-slate-200 text-slate-600 py-3 rounded-[14px] font-bold text-[11px] uppercase tracking-wider shadow-sm hover:bg-slate-50 active:scale-95 transition-all">Impor Excel</button>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <button onClick={() => backupJSON(records, vehicles, showToast)} className="bg-slate-800 text-white py-3 rounded-[14px] font-bold text-[11px] uppercase tracking-wider hover:bg-slate-900 active:scale-95 transition-all shadow-md">Backup JSON</button>
                    <button onClick={() => triggerFileImport('.json')} className="bg-white border border-slate-200 text-slate-600 py-3 rounded-[14px] font-bold text-[11px] uppercase tracking-wider shadow-sm hover:bg-slate-50 active:scale-95 transition-all">Restore JSON</button>
                </div>
                <input type="file" ref={fileInputRef} className="hidden" onChange={handleFileImport} />
            </div>
        </div>
      </div>

      {/* Reminder / Notification Modal */}
      <div className={`fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-end md:items-center justify-center transition-all duration-300 ${isReminderModalOpen ? 'opacity-100 visible' : 'opacity-0 invisible'}`}>
        <div className={`bg-white rounded-t-3xl md:rounded-3xl shadow-2xl border border-slate-100 w-full max-w-sm overflow-hidden transform transition-all duration-300 ${isReminderModalOpen ? 'translate-y-0 scale-100' : 'translate-y-12 md:translate-y-0 md:scale-95'}`}>
            <div className="p-5 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                <h3 className="text-lg font-black text-slate-800 tracking-tight flex items-center gap-2">
                    <Bell size={20} className={theme.text} strokeWidth={2.5} />
                    Sistem Pengingat
                </h3>
                <button onClick={() => setIsReminderModalOpen(false)} className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-200/50 rounded-xl transition-all"><X size={20} strokeWidth={2.5}/></button>
            </div>
            <div className="p-6 space-y-6">
                <div className="space-y-4">
                    {/* Browser Notification Status */}
                    <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 space-y-3">
                        <div className="flex items-center justify-between">
                            <span className="text-xs font-bold text-slate-500 uppercase tracking-widest">Status Notifikasi Browser</span>
                            <span className={`px-2.5 py-1 rounded-full text-[10px] font-black uppercase tracking-wider ${
                                notificationPermission === 'granted' 
                                    ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' 
                                    : notificationPermission === 'denied'
                                    ? 'bg-red-50 text-red-700 border border-red-100'
                                    : 'bg-amber-50 text-amber-700 border border-amber-100'
                            }`}>
                                {notificationPermission === 'granted' ? 'Aktif' : notificationPermission === 'denied' ? 'Ditolak' : 'Belum Aktif'}
                            </span>
                        </div>
                        <p className="text-xs text-slate-500 leading-relaxed">
                            Mengirim pengingat langsung ke sistem operasi/browser Anda saat masa pakai komponen hampir habis.
                        </p>
                        {notificationPermission !== 'granted' ? (
                            <button onClick={requestNotificationPermission} className={`w-full py-2.5 ${theme.bgBase} ${theme.bgHover} text-white font-bold rounded-xl text-xs uppercase tracking-wider shadow-sm transition-all active:scale-95 mt-1`}>
                                Aktifkan Notifikasi
                            </button>
                        ) : (
                            <button onClick={() => checkAndNotifyService(true)} className="w-full py-2.5 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 font-bold rounded-xl text-xs uppercase tracking-wider shadow-sm transition-all active:scale-95 mt-1">
                                Kirim Tes Notifikasi
                            </button>
                        )}
                    </div>

                    {/* Notification threshold selection */}
                    <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest px-1">Ambang Batas Pengingat</label>
                        <select value={reminderThreshold} onChange={e => setReminderThreshold(parseInt(e.target.value))} className={`w-full px-4 py-3 bg-white border border-slate-200 ${theme.borderFocus} shadow-sm rounded-xl text-sm font-semibold ${theme.ringFocus} outline-none text-slate-800 transition-all duration-300`}>
                            <option value="200">200 KM Sebelum Jatuh Tempo</option>
                            <option value="500">500 KM Sebelum Jatuh Tempo</option>
                            <option value="1000">1000 KM Sebelum Jatuh Tempo</option>
                        </select>
                        <p className="text-[10px] text-slate-400 px-1 italic">
                            Anda akan menerima notifikasi/toast ketika sisa umur komponen berada di bawah jarak ini.
                        </p>
                    </div>
                </div>

                <button onClick={() => { checkAndNotifyService(true); setIsReminderModalOpen(false); }} className={`w-full ${theme.bgBase} ${theme.bgHover} text-white font-bold py-3.5 rounded-xl shadow-lg ${theme.shadowHover} active:scale-95 transition-all duration-300 text-xs uppercase tracking-wider`}>
                    Periksa & Simpan Pengaturan
                </button>
            </div>
        </div>
      </div>

      {/* Confirmation Dialog */}
      <div className={`fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-[100] flex items-end md:items-center justify-center transition-all duration-300 ${confirmDialog.isOpen ? 'opacity-100 visible' : 'opacity-0 invisible'}`}>
        <div className={`bg-white rounded-t-3xl md:rounded-3xl shadow-2xl border border-slate-100 w-full max-w-sm overflow-hidden transform transition-all duration-300 ${confirmDialog.isOpen ? 'translate-y-0 scale-100' : 'translate-y-12 md:translate-y-0 md:scale-95'}`}>
            <div className="p-5 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                <h3 className="text-lg font-black text-slate-800 tracking-tight">{confirmDialog.title}</h3>
                <button onClick={confirmDialog.onCancel} className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-200/50 rounded-xl transition-all"><X size={20} strokeWidth={2.5}/></button>
            </div>
            <div className="p-6 space-y-6">
                <p className="text-sm text-slate-600 leading-relaxed font-medium">
                  {confirmDialog.message}
                </p>
                <div className="flex gap-3 mt-4">
                  <button onClick={confirmDialog.onCancel} className="flex-1 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 font-bold py-3.5 rounded-[16px] transition-all active:scale-95 shadow-sm">Batal</button>
                  <button onClick={confirmDialog.onConfirm} className={`flex-1 ${theme.bgBase} ${theme.bgHover} text-white font-bold py-3.5 rounded-[16px] shadow-md ${theme.shadowHover} active:scale-95 transition-all`}>Ya, Lanjutkan</button>
                </div>
            </div>
        </div>
      </div>

    </div>
  );
}
