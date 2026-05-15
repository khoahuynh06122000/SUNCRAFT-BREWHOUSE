/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect, ReactNode, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  LayoutDashboard, 
  PlusCircle, 
  MinusCircle, 
  Package, 
  Users, 
  History, 
  Search, 
  Filter,
  ArrowUpRight,
  ArrowDownLeft,
  Calendar,
  User,
  MoreVertical,
  Download,
  LogOut,
  ChevronRight,
  ChevronDown,
  TrendingUp,
  TrendingDown,
  AlertCircle,
  Info,
  Menu,
  X,
  CreditCard,
  Droplet,
  CheckCircle,
  Truck,
  Camera,
  Image as ImageIcon,
  FileSpreadsheet,
  Layers,
  FileText,
  AlertTriangle,
  DollarSign,
  HandCoins,
  BarChart3,
  Trash2,
  RefreshCw,
  ShieldCheck,
  Package2,
  PackageSearch,
  FileUp,
  Beer,
  Sun
} from 'lucide-react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  Legend, 
  ResponsiveContainer, 
  PieChart, 
  Pie, 
  Cell,
  LineChart,
  Line,
  AreaChart,
  Area
} from 'recharts';

import { 
  isValid,
  parse,
  startOfDay, 
  endOfDay, 
  startOfWeek, 
  endOfWeek, 
  startOfMonth, 
  endOfMonth, 
  startOfYear,
  endOfYear,
  isWithinInterval,
  parseISO,
  format,
  differenceInDays,
  addDays,
  subDays,
  addWeeks,
  subWeeks,
  addMonths,
  subMonths,
  addYears,
  subYears
} from 'date-fns';
import { vi } from 'date-fns/locale';
import { 
  Transaction, 
  Product, 
  Partner, 
  InventoryItem, 
  TransactionType, 
  Category,
  BatchInfo,
  RevenueRecord
} from './types';
import { 
  INITIAL_PRODUCTS, 
  INITIAL_PARTNERS, 
  INITIAL_TRANSACTIONS 
} from './constants';
import { cn, formatDate, formatNumber } from './lib/utils';
import { 
  db, 
  auth, 
  googleProvider, 
  signInWithPopup, 
  signOut, 
  onAuthStateChanged,
  signInAnonymously,
  updateProfile,
  collection,
  doc,
  setDoc,
  updateDoc,
  deleteDoc,
  onSnapshot,
  query,
  orderBy,
  getDocFromServer,
  writeBatch
} from './firebase';

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData?.map(provider => ({
        providerId: provider.providerId,
        email: provider.email,
      })) || []
    },
    operationType,
    path
  }
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  
  // Instead of throwing which crashes the component, we'll alert and log
  const errorMessage = error instanceof Error ? error.message : String(error);
  if (errorMessage.includes('Insufficient permissions')) {
    alert(`Lỗi phân quyền: Anh chưa có quyền thực hiện thao tác ${operationType} trên ${path || 'dữ liệu'}.`);
  }
}

// --- Components ---

const parseDateSafe = (dateStr: any): Date => {
  if (!dateStr) return new Date();
  if (dateStr instanceof Date && !isNaN(dateStr.getTime())) return dateStr;
  
  const trimmed = String(dateStr).trim();
  
  // Try ISO first
  const isoDate = parseISO(trimmed);
  if (isValid(isoDate)) return isoDate;
  
  // Common formats in Vietnam
  const formats = ['dd/MM/yyyy', 'dd.MM.yyyy', 'dd-MM-yyyy', 'yyyy-MM-dd', 'dd/MM/yy', 'd/M/yyyy'];
  for (const fmt of formats) {
    const parsedDate = parse(trimmed, fmt, new Date());
    if (isValid(parsedDate)) return parsedDate;
  }
  
  const fallback = new Date(trimmed);
  return isValid(fallback) ? fallback : new Date();
};

const parseExcelDate = (val: any): string => {
  if (!val) return new Date().toISOString();
  
  // If it's a number (Excel date serial)
  if (typeof val === 'number') {
    // Excel dates are number of days since Dec 30, 1899
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    date.setHours(0, 0, 0, 0);
    return date.toISOString();
  }
  
  const date = parseDateSafe(val);
  date.setHours(0, 0, 0, 0);
  return date.toISOString();
};

const formatDisplayDate = (dateStr: string) => {
  try {
    const date = parseISO(dateStr);
    if (!isValid(date)) {
      // Try to return as-is if it looks like a date string already
      return dateStr;
    }
    return format(date, 'dd/MM/yyyy');
  } catch {
    return dateStr;
  }
};

const Card = ({ children, className, title, noPadding }: { children: ReactNode; className?: string; title?: string; noPadding?: boolean; key?: any }) => (
  <div 
    className={cn("bg-white/80 backdrop-blur-md border border-white/40 rounded-[20px] sm:rounded-[28px] overflow-hidden premium-shadow transition-none", className)}
  >
    {title && (
      <div className="px-4 py-2.5 sm:px-8 sm:py-5 border-b border-slate-50/50 flex items-center justify-between bg-slate-50/20">
        <h3 className="text-[9px] sm:text-xs font-black text-slate-400 uppercase tracking-[0.2em]">{title}</h3>
      </div>
    )}
    <div className={noPadding ? "" : "px-4 py-4 sm:px-8 sm:py-6"}>{children}</div>
  </div>
);

const StatCard = ({ title, value, unit, icon: Icon, color = 'primary', subtitle, target, trend, chartData }: any) => (
  <Card className="relative group overflow-hidden border-none ring-1 ring-slate-100/50 pb-0">
    <div className="flex items-center justify-between mb-2 sm:mb-4">
      <div className={cn(
        "w-9 h-9 sm:w-12 sm:h-12 rounded-[10px] sm:rounded-[16px] flex items-center justify-center transition-all duration-500 group-hover:scale-110 group-hover:rotate-6 shadow-sm",
        color === 'primary' ? "bg-slate-900 text-white shadow-slate-200" : 
        color === 'green' ? "bg-emerald-500 text-white shadow-emerald-100" :
        color === 'amber' ? "bg-amber-500 text-white shadow-amber-100" :
        "bg-rose-500 text-white shadow-rose-100"
      )}>
        {Icon && <Icon className="w-4 h-4 sm:w-6 sm:h-6" />}
      </div>
      {trend && (
        <div className={cn(
          "px-2 py-0.5 sm:px-2.5 sm:py-1 rounded-full text-[8px] sm:text-[10px] font-black tracking-wider flex items-center gap-1 shadow-sm border",
          trend.startsWith('+') ? "bg-emerald-50 border-emerald-100 text-emerald-600" : "bg-rose-50 border-rose-100 text-rose-600"
        )}>
          {trend.startsWith('+') ? <TrendingUp className="w-2.5 h-2.5 sm:w-3 sm:h-3" /> : <TrendingDown className="w-2.5 h-2.5 sm:w-3 sm:h-3" />}
          {trend}
        </div>
      )}
    </div>
    
    <div className="flex flex-col mb-2 sm:mb-4">
      <p className="text-[8px] sm:text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] sm:tracking-[0.25em] mb-0.5">{title}</p>
      <h4 className="text-lg sm:text-3xl font-black text-slate-900 tracking-tight flex items-baseline gap-1.5 sm:gap-2">
        <span className="font-mono leading-none">{value}</span>
        {unit && <span className="text-[8px] sm:text-xs font-black text-slate-400 uppercase tracking-widest">{unit}</span>}
      </h4>
      {target && <p className="text-[8px] sm:text-[10px] font-bold text-slate-400 mt-0.5 sm:mt-1 uppercase tracking-wider uppercase tracking-wider">Mục tiêu: {target}</p>}
      {subtitle && <p className="text-[9px] sm:text-xs font-bold text-slate-500 mt-1.5 sm:mt-3 italic flex items-center gap-1 sm:gap-1.5 font-sans">
        {!trend && <TrendingUp className="w-3 h-3 sm:w-4 sm:h-4 text-emerald-500" />}
        {subtitle}
      </p>}
    </div>

    {chartData && (
      <div className="h-16 w-full -mx-8 mt-2 overflow-hidden">
        <ResponsiveContainer width="100%" height="100%">
          <AreaChart data={chartData}>
            <defs>
              <linearGradient id={`gradient-${color}`} x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor={color === 'primary' ? '#0f172a' : color === 'green' ? '#10b981' : color === 'amber' ? '#f59e0b' : '#f43f5e'} stopOpacity={0.2}/>
                <stop offset="95%" stopColor={color === 'primary' ? '#0f172a' : color === 'green' ? '#10b981' : color === 'amber' ? '#f59e0b' : '#f43f5e'} stopOpacity={0}/>
              </linearGradient>
            </defs>
            <Area 
              type="monotone" 
              dataKey="value" 
              stroke={color === 'primary' ? '#0f172a' : color === 'green' ? '#10b981' : color === 'amber' ? '#f59e0b' : '#f43f5e'} 
              strokeWidth={2}
              fillOpacity={1} 
              fill={`url(#gradient-${color})`} 
            />
          </AreaChart>
        </ResponsiveContainer>
      </div>
    )}

    <div className={cn(
      "absolute bottom-0 right-0 w-32 h-32 blur-[60px] opacity-10 rounded-full -mr-12 -mb-12 transition-all duration-700 group-hover:opacity-20 group-hover:scale-125",
      color === 'primary' ? "bg-primary" : color === 'green' ? "bg-emerald-500" : color === 'amber' ? "bg-amber-500" : "bg-rose-500"
    )} />
  </Card>
);

const Button = ({ children, className, variant = 'primary', ...props }: any) => {
  const variants = {
    primary: "bg-primary text-white hover:bg-primary-dark shadow-lg shadow-primary/25 active:scale-[0.97]",
    secondary: "bg-slate-100 text-slate-700 hover:bg-slate-200 active:scale-[0.97]",
    outline: "border border-slate-200 text-slate-600 hover:bg-slate-50 active:scale-[0.97]",
    danger: "bg-rose-500 text-white hover:bg-rose-600 shadow-lg shadow-rose-200 active:scale-[0.97]",
    ghost: "text-slate-500 hover:bg-slate-100 hover:text-slate-900"
  };
  return (
    <button 
      className={cn(
        "px-4 py-2.5 sm:px-6 sm:py-3.5 rounded-lg sm:rounded-xl font-black transition-all duration-200 flex items-center justify-center gap-2 text-sm sm:text-base disabled:opacity-50 disabled:cursor-not-allowed",
        variants[variant as keyof typeof variants],
        className
      )} 
      {...props}
    >
      {children}
    </button>
  );
};

const Input = ({ label, ...props }: any) => (
  <div className="space-y-1.5 sm:space-y-2 w-full">
    {label && <label className="text-[10px] sm:text-xs font-black text-slate-500 uppercase tracking-widest ml-1">{label}</label>}
    <input 
      className="w-full px-4 py-3 sm:px-5 sm:py-4 bg-slate-50/50 border border-slate-200 rounded-lg sm:rounded-xl text-sm sm:text-base font-bold focus:ring-4 focus:ring-primary/5 focus:border-primary focus:bg-white transition-all outline-none placeholder:text-slate-400" 
      {...props} 
    />
  </div>
);

const Select = ({ label, options, ...props }: any) => (
  <div className="space-y-1.5 sm:space-y-2 w-full">
    {label && <label className="text-[10px] sm:text-xs font-black text-slate-500 uppercase tracking-widest ml-1">{label}</label>}
    <div className="relative">
      <select 
        className="w-full px-4 py-3 sm:px-5 sm:py-4 bg-slate-50/50 border border-slate-200 rounded-lg sm:rounded-xl text-sm sm:text-base font-bold focus:ring-4 focus:ring-primary/5 focus:border-primary focus:bg-white transition-all outline-none appearance-none" 
        {...props}
      >
        {options.map((opt: any) => (
          <option key={opt.value} value={opt.value}>{opt.label}</option>
        ))}
      </select>
      <div className="absolute right-4 sm:right-5 top-1/2 -translate-y-1/2 pointer-events-none opacity-40">
        <ChevronRight className="w-4 h-4 sm:w-5 sm:h-5 rotate-90" />
      </div>
    </div>
  </div>
);

// --- Main Application ---

export default function App() {
  const [user, setUser] = useState<string | null>(null);
  const [userEmail, setUserEmail] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [notification, setNotification] = useState<{ message: string; type: 'success' | 'error' } | null>(null);

  const showNotification = (message: string, type: 'success' | 'error' = 'success') => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 3000);
  };
  const [activeTab, setActiveTab] = useState('dashboard');
  const [reportSubTab, setReportSubTab] = useState<'summary' | 'in' | 'out'>('summary');
  const [timeFilter, setTimeFilter] = useState<'day' | 'week' | 'month' | 'year' | 'all'>('all');
  const [filterBaseDate, setFilterBaseDate] = useState<Date>(new Date());
  const dateInputRef = useRef<HTMLInputElement>(null);
  const [historySearchQuery, setHistorySearchQuery] = useState('');
  const [batchSearchQuery, setBatchSearchQuery] = useState('');
  const [galleryFilter, setGalleryFilter] = useState<'IN' | 'OUT'>('OUT');
  const [selectedGalleryImage, setSelectedGalleryImage] = useState<Transaction | null>(null);
  const [selectedInventoryProduct, setSelectedInventoryProduct] = useState<string | null>(null);
  
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [partners, setPartners] = useState<Partner[]>([]);
  const [revenueData, setRevenueData] = useState<RevenueRecord[]>([]);
  const [expandedInvoices, setExpandedInvoices] = useState<string[]>([]);
  const [revenuePartnerSearch, setRevenuePartnerSearch] = useState('');
  const [galleryMonth, setGalleryMonth] = useState<string>('all');
  const [gallerySearchQuery, setGallerySearchQuery] = useState<string>('');

  const [products] = useState<Product[]>(INITIAL_PRODUCTS);
  const [showGuestNameModal, setShowGuestNameModal] = useState(false);
  const [guestNameInput, setGuestNameInput] = useState('');

  // AUTH OBSERVER
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (firebaseUser) => {
      if (firebaseUser) {
        if (firebaseUser.isAnonymous) {
          setUser(firebaseUser.displayName || 'Guest User');
          setUserEmail(null);
        } else {
          setUser(firebaseUser.displayName || firebaseUser.email || 'Expert');
          setUserEmail(firebaseUser.email);
        }
      } else {
        setUser(null);
        setUserEmail(null);
      }
      setLoading(false);
    });
    return () => unsubscribe();
  }, []);

  // FIRESTORE SYNC
  useEffect(() => {
    if (!user) return;

    // Test connection as required by security guidelines
    const testConnection = async () => {
      try {
        await getDocFromServer(doc(db, 'test', 'connection'));
      } catch (error) {
        if (error instanceof Error && error.message.includes('the client is offline')) {
          console.error("Please check your Firebase configuration.");
        }
      }
    };
    testConnection();

    // Sync Transactions
    const qTransactions = query(collection(db, 'transactions'), orderBy('date', 'desc'));
    const unsubTransactions = onSnapshot(qTransactions, (snapshot) => {
      const data = snapshot.docs.map(doc => ({ ...doc.data(), id: doc.id } as Transaction));
      setTransactions(data.length > 0 ? data : INITIAL_TRANSACTIONS);
    }, (error) => {
      handleFirestoreError(error, OperationType.GET, 'transactions');
    });

    // Sync Partners
    const qPartners = query(collection(db, 'partners'));
    const unsubPartners = onSnapshot(qPartners, (snapshot) => {
      const data = snapshot.docs.map(doc => ({ ...doc.data(), id: doc.id } as Partner));
      setPartners(data.length > 0 ? data : INITIAL_PARTNERS);
    }, (error) => {
      handleFirestoreError(error, OperationType.GET, 'partners');
    });

    // Sync Revenue
    const qRevenue = query(collection(db, 'revenue'));
    const unsubRevenue = onSnapshot(qRevenue, (snapshot) => {
      const data = snapshot.docs.map(doc => ({ ...doc.data(), id: doc.id } as RevenueRecord));
      setRevenueData(data);
    }, (error) => {
      handleFirestoreError(error, OperationType.GET, 'revenue');
    });

    return () => {
      unsubTransactions();
      unsubPartners();
      unsubRevenue();
    };
  }, [user]);

  const [isAuthenticating, setIsAuthenticating] = useState(false);

  const handleLogin = async () => {
    if (isAuthenticating) return;
    setIsAuthenticating(true);
    try {
      await signInWithPopup(auth, googleProvider);
    } catch (error: any) {
      console.error("Login detail:", error);
      const errorCode = error?.code || "";
      const errorMessage = error?.message || "";
      
      // Kiểm tra các lỗi người dùng chủ động đóng hoặc hủy
      if (errorCode === 'auth/popup-closed-by-user' || errorCode === 'auth/cancelled-popup-request') {
        return;
      }
      
      // Kiểm tra lỗi tên miền chưa được cấp phép
      if (errorCode === 'auth/unauthorized-domain' || errorMessage.includes('unauthorized-domain')) {
        alert(
          "LỖI PHÂN QUYỀN: Tên miền này chưa được cấp phép trong Firebase.\n\n" +
          "Anh hãy làm theo bước này nhé:\n" +
          "1. Copy tên miền này: banabrewhouse.vercel.app\n" +
          "2. Vào Firebase Console -> Authentication -> Settings.\n" +
          "3. Thêm nó vào mục 'Authorized domains' là xong ạ!"
        );
      } else {
        alert(`Đăng nhập thất bại. Vui lòng thử lại. (Mã lỗi: ${errorCode || 'Unknown'})`);
      }
    } finally {
      setIsAuthenticating(false);
    }
  };

  const handleGuestLogin = async () => {
    if (!guestNameInput.trim()) {
      setShowGuestNameModal(true);
      return;
    }

    try {
      const userCredential = await signInAnonymously(auth);
      if (userCredential.user) {
        await updateProfile(userCredential.user, {
          displayName: guestNameInput.trim()
        });
        // Force refresh user state
        setUser(guestNameInput.trim());
      }
      setShowGuestNameModal(false);
      setGuestNameInput('');
    } catch (error: any) {
      console.error("Guest login failed:", error);
      if (error.code === 'auth/admin-restricted-operation') {
        alert(
          "LỖI HỆ THỐNG: Tính năng 'Đăng nhập ẩn danh' chưa được bật trên Firebase Console.\n\n" +
          "Cách sửa:\n" +
          "1. Truy cập: https://console.firebase.google.com/project/gen-lang-client-0780471401/authentication/providers\n" +
          "2. Chọn 'Anonymous' và nhấn 'Enable'.\n" +
          "3. Quan trọng: Nhấn nút 'SAVE' màu xanh để lưu lại.\n\n" +
          "Sau khi làm xong, anh hãy nhấn lại nút 'Khách' nhé!"
        );
      } else {
        alert("Truy cập khách thất bại. Vui lòng thử lại.");
      }
    }
  };

  const handleHardReset = async () => {
    if (!isOwner) return;
    
    const confirmReset = window.confirm(
      "TIÊU ĐIỂM: Tin sẽ giúp anh làm sạch toàn bộ dữ liệu Giao dịch và Doanh thu.\n\nMục Đối tác và Danh mục sản phẩm sẽ được giữ lại theo yêu cầu của anh.\n\nAnh có chắc chắn muốn dọn dẹp không?"
    );
    
    if (!confirmReset) return;
    
    try {
      setLoading(true);
      const batch = writeBatch(db);
      let count = 0;
      
      // Delete all Transactions
      transactions.forEach(t => {
        if (t.id) {
          batch.delete(doc(db, 'transactions', t.id));
          count++;
        }
      });
      
      // Delete all Revenue Data
      revenueData.forEach(r => {
        if (r.id) {
          batch.delete(doc(db, 'revenue', r.id));
          count++;
        }
      });

      if (count === 0) {
        showNotification("Hệ thống hiện tại đã ở trạng thái sạch.");
        setLoading(false);
        return;
      }

      await batch.commit();
      setLoading(false);
      showNotification('Hệ thống cập nhật data thành công');
    } catch (error) {
      console.error("Hard Reset Error:", error);
      const errorMsg = error instanceof Error ? error.message : String(error);
      if (errorMsg.includes('unavailable')) {
        alert("LỖI KẾT NỐI: Tin không thể kết nối tới máy chủ Firestore. Anh hãy thử tải lại trang (F5) hoặc kiểm tra lại mạng nhé.");
      } else {
        alert("Có lỗi xảy ra khi dọn dẹp. Anh kiểm tra lại xem đã đăng nhập đúng tài khoản " + userEmail + " chưa nhé.");
      }
      setLoading(false);
    }
  };

  const handleRestorePartners = async () => {
    if (!isOwner) return;
    
    if (!window.confirm("Tin Tin sẽ giúp anh khôi phục lại toàn bộ danh sách Đối tác (SKB, APC, BNG, Capella...) vào hệ thống. Anh đồng ý chứ?")) return;

    setLoading(true);
    try {
      const batch = writeBatch(db);
      INITIAL_PARTNERS.forEach(p => {
        batch.set(doc(db, 'partners', p.id), p);
      });
      await batch.commit();
      setLoading(false);
      showNotification('Khôi phục danh sách Đối tác thành công');
    } catch (error) {
      console.error("Restore Partners Error:", error);
      alert("Có lỗi khi khôi phục đối tác. Anh kiểm tra lại kết nối mạng hoặc phân quyền nhé.");
    } finally {
      setLoading(false);
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
    } catch (error) {
      console.error("Logout failed:", error);
    }
  };

  const sidebarOpenRef = useRef(true);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const revenueInputRef = useRef<HTMLInputElement>(null);
  const handleExportReportToExcel = () => {
    let exportData: any[] = [];
    let sheetName = "Report";

    if (reportSubTab === 'summary') {
      exportData = flowSummary.map(item => ({
        'Sản phẩm': item.productName,
        'Mã SP': item.id,
        'Quy cách': item.category,
        'Đơn vị': item.unit,
        'Tồn đầu': item.openingStock,
        'Tổng Nhập': item.in,
        'Tổng Xuất': item.out,
        'Tồn cuối': item.closingStock,
        'Giá trị Nhập': item.inValue,
        'Giá trị Xuất': item.outValue
      }));
      sheetName = "Tong_Hop_Kho";
    } else {
      const typeLabel = reportSubTab === 'in' ? 'Nhap' : 'Xuat';
      const targets = filteredTransactionsForReport.filter(t => 
        reportSubTab === 'in' ? (t.type === 'IN' || t.type === 'OPENING') : t.type === 'OUT'
      );

      if (targets.length === 0) {
        showNotification(`Không có dữ liệu ${typeLabel} để xuất.`, 'error');
        return;
      }

      exportData = targets.map(t => ({
        'Ngày': format(parseISO(t.date), 'dd/MM/yyyy'),
        'Loại': t.type,
        'Sản phẩm': t.productName,
        'Số lượng': t.quantity,
        'Đối tác': t.partnerName,
        'Ghi chú': t.notes || '',
        'Số lô': t.batchNumber || ''
      }));
      sheetName = `Bao_cao_${typeLabel}`;
    }

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    
    const fileName = `${sheetName}_${format(new Date(), 'ddMMyyyy_HHmm')}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  const clearAllRevenueData = async () => {
    if (!isOwner) {
      alert("Chỉ anh Khoa mới có quyền xóa sạch dữ liệu doanh thu ạ!");
      return;
    }
    if (!window.confirm('CẢNH BÁO: Hành động này sẽ XÓA TOÀN BỘ dữ liệu doanh thu đang có trên hệ thống. Bạn có chắc chắn muốn thực hiện không?')) return;
    
    setLoading(true);
    try {
      const deletePromises = revenueData.map(r => deleteDoc(doc(db, 'revenue', r.id)));
      
      if (deletePromises.length > 0) {
        await Promise.all(deletePromises);
        setLoading(false);
        showNotification(`Đã xóa sạch toàn bộ ${deletePromises.length} dòng dữ liệu doanh thu.`);
      } else {
        setLoading(false);
        showNotification('Hệ thống hiện không có dữ liệu doanh thu nào để xóa.');
      }
    } catch (err) {
      console.error(err);
      setLoading(false);
      showNotification('Lỗi khi xóa dữ liệu doanh thu.', 'error');
    }
  };

  const downloadRevenueTemplate = () => {
    const templateData = [{
      'Ngày hóa đơn': '14.05.2026',
      'Đơn vị': 'BNC',
      'Tên hàng hóa': 'Bia Golden Bridge Helles Lager',
      'Đơn vị tính': 'LIT',
      'Số lượng': 13805.10,
      'Đơn giá': 30000,
      'Thành tiền': 414153000,
      'VAT': 41415300,
      'Thành tiền sau thuế': 455568300,
      'Số hóa đơn': 'C26TKB#00000093'
    }, {
      'Ngày hóa đơn': '14.05.2026',
      'Đơn vị': 'BNG',
      'Tên hàng hóa': 'Bia Wings Dark Lager 330ml',
      'Đơn vị tính': 'LON',
      'Số lượng': 72.00,
      'Đơn giá': 14000,
      'Thành tiền': 1008000,
      'VAT': 100800,
      'Thành tiền sau thuế': 1108800,
      'Số hóa đơn': 'C26TKB#00000094'
    }];
    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Mau_Doanh_Thu");
    XLSX.writeFile(wb, "Mau_Nap_Doanh_Thu_TinTin.xlsx");
    showNotification("Đã tải file mẫu thành công. Anh dùng file này để nạp nhé!");
  };

  const handleRevenueReportUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Security check: Removed as requested, now available to all users
    const isAuthorized = true;

    if (!isAuthorized) {
      alert('Chỉ quản trị viên hoặc Khoahuynh mới có quyền nạp dữ liệu hệ thống.');
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const bstr = event.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws) as any[];

        if (data.length === 0) {
          alert('File không có dữ liệu hoặc định dạng không đúng.');
          return;
        }

        // Helper to generate a unique signature for a record
        const generateSignature = (r: any) => {
          // Normalize date to YYYY-MM-DD for signature to avoid time mismatches
          let dStr = 'NO_DATE';
          try {
            if (r.date) {
              const d = new Date(r.date);
              if (!isNaN(d.getTime())) {
                dStr = d.toISOString().split('T')[0];
              } else if (typeof r.date === 'string') {
                dStr = r.date.split('T')[0];
              }
            }
          } catch(e) {}
          
          return `${dStr}|${r.invoiceNumber || 'NO_INV'}|${r.productName}|${r.quantity}|${r.unitPrice}|${r.partnerName}`.toLowerCase().replace(/\s+/g, '');
        };

        // Step 1: Identify all unique Invoice Numbers in the incoming file
        const incomingInvoices = new Set<string>();
        data.forEach(row => {
          const inv = String(row['Số hóa đơn'] || row['Số HĐ'] || row['Invoice No'] || row['Số chứng từ'] || row['Số CT'] || '').trim();
          if (inv && inv !== '') incomingInvoices.add(inv);
        });

        // Step 2: Identify existing invoice numbers in DB to prevent duplicates
        const dbInvoices = new Set<string>();
        revenueData.forEach(r => {
          if (r.invoiceNumber) dbInvoices.add(r.invoiceNumber.trim());
        });

        const newRecordsPromises: Promise<any>[] = [];
        let newAddedCount = 0;
        let skippedCount = 0;

        data.forEach((row, idx) => {
          const invNumber = String(row['Số hóa đơn'] || row['Số HĐ'] || row['Invoice No'] || row['Số chứng từ'] || row['Số CT'] || '').trim();
          
          // If this invoice already exists in the database, skip it entirely
          if (invNumber && dbInvoices.has(invNumber)) {
            skippedCount++;
            return;
          }

          const dateVal = row['Ngày hóa đơn'] || row['Ngày xuất hoá đơn'] || row['Ngày hóa đơn (ngày nhận)'] || row['Ngày giao bia'] || row['Ngày'] || row['Date'] || row['Ngày chứng từ'];
          const excelPartnerName = String(
            row['Đơn vị thụ hưởng'] || 
            row['Đơn vị'] || 
            row['Don vi'] ||
            row['Mã đơn vị'] ||
            row['Ma don vi'] ||
            row['Tên đơn vị'] ||
            row['Ten don vi'] ||
            row['Partner'] || 
            row['Khách hàng'] || 
            row['Khach hang'] ||
            row['Mã KH'] || 
            row['Ma KH'] ||
            row['Đối tác'] ||
            row['Doi tac'] ||
            row['Tên đối tác'] ||
            row['Ten doi tac'] ||
            row['Customer'] ||
            '—'
          ).trim();
          
          const officialPartner = partners.find(p => 
            p.name.toLowerCase() === excelPartnerName.toLowerCase() || 
            p.id.toLowerCase() === excelPartnerName.toLowerCase()
          );
          
          const productName = row['Tên hàng hóa'] || row['Sản phẩm'] || row['Product'] || row['Tên sản phẩm'] || row['Hàng hóa/Dịch vụ'] || 'Sản phẩm không tên';
          const qty = Number(row['Số lượng'] || row['Quantity'] || row['SL'] || 0);

          const record = {
            id: '', 
            date: parseExcelDate(dateVal),
            productName: productName,
            materialCode: String(row['Mã vật tư'] || row['Mã hàng'] || row['Material Code'] || row['Mã sản phẩm'] || row['Mã số'] || ''),
            unit: String(row['Đơn vị tính'] || row['ĐVT'] || row['Unit'] || row['Đơn vị'] || ''),
            quantity: qty,
            unitPrice: Number(row['SKB - TLD'] || row['ĐG TLD'] || row['Đơn giá'] || row['Price'] || row['Unit Price'] || row['ĐG'] || 0),
            totalAmount: Number(row['Thành tiền sau thuế'] || row['Thành tiền'] || row['Total'] || row['Revenue Amount'] || row['Tổng tiền'] || 0),
            vatAmount: Number(row['VAT'] || row['Thuế GTGT'] || 0),
            invoiceNumber: invNumber,
            deptCode: String(row['Mã BP'] || row['Bộ phận'] || row['Mã phòng ban'] || ''),
            partnerName: officialPartner ? officialPartner.name : excelPartnerName,
            partnerId: officialPartner?.id || ''
          };

          const deterministicId = `rev-${record.invoiceNumber || 'INV'}-${idx}-${Date.now()}`;
          record.id = deterministicId;
          newRecordsPromises.push(setDoc(doc(db, 'revenue', deterministicId), record));

          newAddedCount++;
        });

        if (newRecordsPromises.length === 0) {
          if (skippedCount > 0) {
            showNotification(`Tất cả ${skippedCount} dòng đã tồn tại trên hệ thống. Không có dữ liệu mới được nạp.`, 'error');
          } else {
            showNotification('File không có dữ liệu hợp lệ để xử lý.', 'error');
          }
          setLoading(false);
          return;
        }

        setLoading(true);
        Promise.all(newRecordsPromises).then(() => {
          setActiveTab('revenue-mgmt');
          setLoading(false);
          showNotification(`Đã nạp mới ${newAddedCount} dòng dữ liệu. Bỏ qua ${skippedCount} dòng do trùng Số hóa đơn.`);
          if (revenueInputRef.current) revenueInputRef.current.value = '';
        }).catch(err => {
          console.error(err);
          setLoading(false);
          showNotification('Lỗi khi cập nhật dữ liệu. Vui lòng thử lại.', 'error');
        });
      } catch (err) {
        console.error('Error parsing file:', err);
        showNotification('Có lỗi xảy ra khi nạp dữ liệu lên Cloud.', 'error');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleExportRevenueToExcel = () => {
    if (revenueData.length === 0) {
      showNotification('Không có dữ liệu doanh thu để xuất.', 'error');
      return;
    }

    const exportData = filteredRevenueByTime.map(r => ({
      'Ngày xuất hoá đơn': format(parseISO(r.date), 'dd/MM/yyyy'),
      'Số hóa đơn': r.invoiceNumber,
      'Đơn vị thụ hưởng': r.partnerName,
      'Mã vật tư': r.materialCode,
      'Tên hàng hóa': r.productName,
      'Đơn vị tính': r.unit,
      'Số lượng': r.quantity,
      'SKB - TLD': r.unitPrice,
      'Thành tiền sau thuế': r.totalAmount,
      'VAT': r.vatAmount,
      'Mã BP': r.deptCode
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Revenue");
    
    const fileName = `Bao_cao_Doanh_thu_${format(new Date(), 'ddMMyyyy_HHmm')}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  // Error Handling Helper
  const handleFirestoreError = (err: any, operation: 'create' | 'update' | 'delete' | 'list' | 'get' | 'write', path: string | null = null) => {
    console.error(`Firestore Error [${operation}]:`, err);
    if (err?.message?.includes('Missing or insufficient permissions')) {
      const errorInfo = {
        error: "Missing or insufficient permissions",
        operationType: operation,
        path: path,
        authInfo: {
          userId: auth.currentUser?.uid || 'unknown',
          email: auth.currentUser?.email || 'unknown',
          emailVerified: auth.currentUser?.emailVerified || false,
          isAnonymous: auth.currentUser?.isAnonymous || false,
          providerInfo: auth.currentUser?.providerData.map(p => ({
            providerId: p.providerId,
            displayName: p.displayName || '',
            email: p.email || ''
          })) || []
        }
      };
      throw new Error(JSON.stringify(errorInfo));
    }
    throw err;
  };

  const handleDeleteSyncData2104 = async () => {
    if (!isOwner) {
      alert("Tin Tin rất tiếc, chỉ anh Khoa mới có quyền thực hiện thao tác 'dọn dẹp' này ạ.");
      return;
    }
    if (!window.confirm("Tin Tin sẽ xóa toàn bộ các giao dịch đã đồng bộ tự động ngày 21/04 để anh tự nạp tay nhé?")) return;
    
    setLoading(true);
    let successCount = 0;
    let failCount = 0;
    
    try {
      const syncIds = [
        'p1', 'p15', 'p2', 'p10', 'p11', 'p12', 'p16', 'p14', 'p4', 'p17', 'p5'
      ].flatMap(pid => [`sync-in-${pid}-20260421`, `sync-out-${pid}-20260421`]);

      for (const id of syncIds) {
        try {
          await deleteDoc(doc(db, 'transactions', id));
          successCount++;
        } catch (err: any) {
          console.error(`Failed to delete ${id}:`, err);
          failCount++;
        }
      }

      if (failCount > 0) {
        if (successCount > 0) {
          showNotification(`Đã dọn dẹp được ${successCount} mục, nhưng có ${failCount} mục thất bại.`, 'error');
        } else {
          handleFirestoreError(new Error("Missing or insufficient permissions"), 'delete', 'transactions/*');
        }
      } else {
        showNotification("Đã dọn dẹp sạch sẽ dữ liệu đồng bộ cũ.");
      }
    } catch (err: any) {
      console.error(err);
      let errorMessage = "Dọn dẹp thất bại ạ.";
      try {
        const parsed = JSON.parse(err.message);
        if (parsed.error === "Missing or insufficient permissions") {
          errorMessage = "Lỗi phân quyền: Tin Tin chưa kịp cập nhật quyền xóa cho tài khoản " + (auth.currentUser?.email || "của anh") + ". Anh đợi 1-2 phút rồi thử lại nhé!";
        }
      } catch {
        errorMessage = "Dọn dẹp thất bại: " + (err instanceof Error ? err.message : String(err));
      }
      alert(errorMessage);
    } finally {
      setLoading(false);
    }
  };

  const handleImportInventoryExcel = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!window.confirm("Bắt đầu nhập dữ liệu Tồn kho từ file Excel?")) return;

    setLoading(true);
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];

        let successCount = 0;
        const syncBy = `User Upload (${user})`;
        const syncDate = new Date().toISOString();

        // Deep clone batches to track allocation through the loop
        let runningBatches = batches.map(b => ({ ...b }));

        for (const row of jsonData) {
          const productName = row['Sản phẩm'] || row['Product'] || '';
          const totalImport = parseFloat(row['Tổng Nhập'] || row['Total Import'] || '0');
          const totalExport = parseFloat(row['Tổng Xuất'] || row['Total Export'] || '0');
          const openingStock = parseFloat(row['Tồn Đầu'] || row['Opening Stock'] || '0');
          const unit = row['Đơn vị'] || row['Unit'] || '';

          if (!productName) continue;

          const product = products.find(p => p.name.toLowerCase().trim() === productName.toLowerCase().trim());
          
          if (product) {
            let finalIn = totalImport + openingStock;
            let finalOut = totalExport;

            if (product.category === 'Lon' && unit.toLowerCase().trim() === 'lon') {
              finalIn = finalIn / (product.conversionFactor || 1);
              finalOut = finalOut / (product.conversionFactor || 1);
            }

            if (finalIn > 0) {
              const importId = `import-in-${product.id}-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`;
              await setDoc(doc(db, 'transactions', importId), {
                id: importId,
                date: syncDate,
                type: 'IN',
                productId: product.id,
                productName: product.name,
                quantity: finalIn,
                partnerId: 'SYSTEM_SYNC',
                partnerName: 'Excel Import',
                category: product.category,
                createdBy: syncBy,
                batchNumber: `IMPORT-${format(new Date(), 'ddMM')}`
              });
            }

            if (finalOut > 0) {
              const allocations = getFIFOAllocations(product.id, finalOut, runningBatches);
              for (const alloc of allocations) {
                const exportId = `import-out-${product.id}-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`;
                await setDoc(doc(db, 'transactions', exportId), {
                  id: exportId,
                  date: syncDate,
                  type: 'OUT',
                  productId: product.id,
                  productName: product.name,
                  quantity: alloc.quantity,
                  partnerId: 'SYSTEM_SYNC',
                  partnerName: 'Excel Import',
                  category: product.category,
                  createdBy: syncBy,
                  batchNumber: alloc.batchNumber
                });
              }
            }
            successCount++;
          }
        }
        setLoading(false);
        showNotification('Hệ thống cập nhật data thành công');
      } catch (err) {
        console.error(err);
        setLoading(false);
        showNotification("Lỗi khi đọc file Excel. Vui lòng kiểm tra định dạng.", "error");
      } finally {
        if (event.target) event.target.value = '';
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // FIFO Helper for auto-allocation
const compressImage = (base64Str: string, maxWidth = 1024, maxHeight = 1024, quality = 0.6): Promise<string> => {
  return new Promise((resolve) => {
    const img = new Image();
    img.src = base64Str;
    img.onload = () => {
      const canvas = document.createElement('canvas');
      let width = img.width;
      let height = img.height;

      if (width > height) {
        if (width > maxWidth) {
          height *= maxWidth / width;
          width = maxWidth;
        }
      } else {
        if (height > maxHeight) {
          width *= maxHeight / height;
          height = maxHeight;
        }
      }

      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      ctx?.drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL('image/jpeg', quality));
    };
    img.onerror = () => resolve(base64Str); // Fallback to original
  });
};

  const getFIFOAllocations = (productId: string, quantity: number, currentBatches: BatchInfo[]) => {
    const allocations: { batchNumber: string, quantity: number }[] = [];
    let remaining = quantity;
    
    // Process batches for this product that have stock
    const relevant = currentBatches
      .filter(b => b.productId === productId && b.stock > 0)
      .sort((a, b) => new Date(a.importDate).getTime() - new Date(b.importDate).getTime());

    for (const b of relevant) {
      if (remaining <= 0) break;
      const taken = Math.min(b.stock, remaining);
      allocations.push({ batchNumber: b.batchNumber, quantity: taken });
      remaining -= taken;
      // Update local copy so subsequent items in the same sync loop see updated state
      b.stock -= taken; 
    }

    if (remaining > 0) {
      allocations.push({ batchNumber: 'VUOT_DINH_MUC', quantity: remaining });
    }
    return allocations;
  };

  // Form states
  interface TransactionItem {
    productId: string;
    quantity: number;
    batchNumber: string;
  }

  const [newTransaction, setNewTransaction] = useState<{
    type: TransactionType;
    partnerId: string;
    notes: string;
    evidencePhotoUrl: string;
    date: string;
    isInTransit: boolean;
    items: TransactionItem[];
  }>({
    type: 'IN',
    partnerId: partners[0]?.id || '',
    notes: '',
    evidencePhotoUrl: '',
    date: format(new Date(), 'yyyy-MM-dd'),
    isInTransit: false,
    items: [{ productId: products[0]?.id || '', quantity: 0, batchNumber: '' }]
  });

  const addTransactionItem = () => {
    setNewTransaction(prev => ({
      ...prev,
      items: [...prev.items, { productId: products[0]?.id || '', quantity: 0, batchNumber: prev.type === 'OPENING' ? 'Tồn 25/4' : '' }]
    }));
  };

  const removeTransactionItem = (index: number) => {
    if (newTransaction.items.length <= 1) return;
    setNewTransaction(prev => ({
      ...prev,
      items: prev.items.filter((_, i) => i !== index)
    }));
  };

  const updateTransactionItem = (index: number, updates: Partial<TransactionItem>) => {
    setNewTransaction(prev => ({
      ...prev,
      items: prev.items.map((item, i) => i === index ? { ...item, ...updates } : item)
    }));
  };

  const [selectedInTransit, setSelectedInTransit] = useState<Transaction | null>(null);
  const [actualReceivedQty, setActualReceivedQty] = useState<number>(0);
  const [lossReason, setLossReason] = useState<string>('');
  const [lossEvidencePhoto, setLossEvidencePhoto] = useState<string>('');
  const [showLossModal, setShowLossModal] = useState(false);

  const [showAddPartner, setShowAddPartner] = useState(false);
  const [partnerFormData, setPartnerFormData] = useState<Omit<Partner, 'id'>>({
    name: '',
    sapCode: '',
    type: 'AGENT',
    phone: '',
    address: ''
  });

  const handleAddPartner = (e: React.FormEvent) => {
    e.preventDefault();
    if (!partnerFormData.name) return;

    const newId = partnerFormData.sapCode || `p-${Math.random().toString(36).substr(2, 6)}`;
    const newPartner: Partner = {
      id: newId,
      ...partnerFormData
    };

    setDoc(doc(db, 'partners', newPartner.id), newPartner).then(() => {
      setShowAddPartner(false);
      setPartnerFormData({ name: '', sapCode: '', type: 'AGENT', phone: '', address: '' });
    }).catch(err => {
      handleFirestoreError(err, OperationType.WRITE, `partners/${newPartner.id}`);
    });
  };

  const handleReportLoss = async () => {
    if (!selectedInTransit || actualReceivedQty < 0) return;
    
    try {
      setLoading(true);
      const diff = selectedInTransit.quantity - actualReceivedQty;
      
      if (diff > 0) {
        // Create loss transaction
        const lossId = `loss-${Date.now()}`;
        const lossTrx: Transaction = {
          ...selectedInTransit,
          id: lossId,
          type: 'LOSS',
          quantity: diff,
          notes: `[HAO HỤT TỪ ĐƠN ${selectedInTransit.id}] Lý do: ${lossReason}`,
          status: 'completed',
          date: new Date().toISOString(),
          evidencePhotoUrl: lossEvidencePhoto || undefined
        };
        await setDoc(doc(db, 'transactions', lossId), lossTrx);

        // Update original transaction
        await updateDoc(doc(db, 'transactions', selectedInTransit.id), {
          quantity: actualReceivedQty,
          status: 'completed',
          notes: `${selectedInTransit.notes || ''} (Đã điều chỉnh từ ${selectedInTransit.quantity} do hao hụt)`.trim(),
          updatedAt: new Date().toISOString(),
          evidencePhotoUrl: lossEvidencePhoto || selectedInTransit.evidencePhotoUrl
        });
      } else {
        // Just confirm with possible new photo
        await updateDoc(doc(db, 'transactions', selectedInTransit.id), {
          status: 'completed',
          updatedAt: new Date().toISOString(),
          evidencePhotoUrl: lossEvidencePhoto || selectedInTransit.evidencePhotoUrl,
          notes: `${selectedInTransit.notes || ''} (Đã xác nhận nhận đủ)`.trim()
        });
      }

      setShowLossModal(false);
      setSelectedInTransit(null);
      setActualReceivedQty(0);
      setLossReason('');
      setLossEvidencePhoto('');
      showNotification('Hệ thống cập nhật data thành công');
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, selectedInTransit ? `transactions/${selectedInTransit.id}` : 'transactions');
    } finally {
      setLoading(false);
    }
  };
  const handleDeletePartner = (id: string) => {
    if (!isOwner) {
      alert("Chỉ anh Khoa mới có quyền xóa đối tác ạ!");
      return;
    }
    if (window.confirm('Bạn có chắc chắn muốn xóa đối tác này không? Dữ liệu giao dịch liên quan sẽ không bị xóa nhưng sẽ mất liên kết tên đối tác chính thức.')) {
      setPartners(partners.filter(p => p.id !== id));
    }
  };

  // Filtering Logic
  const dateRange = useMemo(() => {
    if (timeFilter === 'all') return null;
    
    let start: Date;
    let end: Date;

    if (timeFilter === 'day') {
      start = startOfDay(filterBaseDate);
      end = endOfDay(filterBaseDate);
    } else if (timeFilter === 'week') {
      start = startOfWeek(filterBaseDate, { weekStartsOn: 1 });
      end = endOfWeek(filterBaseDate, { weekStartsOn: 1 });
    } else if (timeFilter === 'month') {
      start = startOfMonth(filterBaseDate);
      end = endOfMonth(filterBaseDate);
    } else {
      start = startOfYear(filterBaseDate);
      end = endOfYear(filterBaseDate);
    }

    return { start, end };
  }, [timeFilter, filterBaseDate]);

  const filteredTransactionsByTime = useMemo(() => {
    if (!dateRange) return transactions;
    
    return transactions.filter(t => {
      const date = parseISO(t.date);
      return isWithinInterval(date, { start: dateRange.start, end: dateRange.end });
    });
  }, [transactions, dateRange]);

  const filteredRevenueByTime = useMemo(() => {
    if (timeFilter === 'all') return revenueData;

    let start: Date;
    let end: Date;

    if (timeFilter === 'day') {
      start = startOfDay(filterBaseDate);
      end = endOfDay(filterBaseDate);
    } else if (timeFilter === 'week') {
      start = startOfWeek(filterBaseDate, { weekStartsOn: 1 });
      end = endOfWeek(filterBaseDate, { weekStartsOn: 1 });
    } else if (timeFilter === 'month') {
      start = startOfMonth(filterBaseDate);
      end = endOfMonth(filterBaseDate);
    } else {
      start = startOfYear(filterBaseDate);
      end = endOfYear(filterBaseDate);
    }

    return revenueData.filter(r => {
      try {
        const date = parseDateSafe(r.date);
        return isWithinInterval(date, { start, end });
      } catch {
        return false;
      }
    });
  }, [revenueData, timeFilter, filterBaseDate]);

  const groupedRevenue = useMemo(() => {
    const groups = new Map<string, {
      invoiceNumber: string;
      date: string;
      partnerName: string;
      totalAmount: number;
      items: RevenueRecord[];
    }>();

    filteredRevenueByTime.forEach(r => {
      // Filter by partner search
      if (revenuePartnerSearch && !r.partnerName.toLowerCase().includes(revenuePartnerSearch.toLowerCase())) {
        return;
      }

      const key = r.invoiceNumber || `N/A-${r.id}`;
      const existing = groups.get(key);

      if (existing) {
        existing.totalAmount += r.totalAmount;
        existing.items.push(r);
      } else {
        groups.set(key, {
          invoiceNumber: r.invoiceNumber || 'N/A',
          date: r.date,
          partnerName: r.partnerName,
          totalAmount: r.totalAmount,
          items: [r]
        });
      }
    });

    return Array.from(groups.values()).sort((a, b) => {
      return parseDateSafe(b.date).getTime() - parseDateSafe(a.date).getTime();
    });
  }, [filteredRevenueByTime, revenuePartnerSearch]);

  const periodLabel = useMemo(() => {
    if (timeFilter === 'all') return 'Tất cả thời gian';
    if (timeFilter === 'day') return format(filterBaseDate, "'Ngày' dd/MM/yyyy", { locale: vi });
    if (timeFilter === 'week') {
      const start = startOfWeek(filterBaseDate, { weekStartsOn: 1 });
      const end = endOfWeek(filterBaseDate, { weekStartsOn: 1 });
      return `Tuần ${format(start, 'dd/MM')} - ${format(end, 'dd/MM/yyyy')}`;
    }
    if (timeFilter === 'month') return format(filterBaseDate, "'Tháng' MM/yyyy", { locale: vi });
    if (timeFilter === 'year') return format(filterBaseDate, "'Năm' yyyy", { locale: vi });
    return '';
  }, [timeFilter, filterBaseDate]);

  const moveFilterDate = (direction: 'prev' | 'next') => {
    if (timeFilter === 'all') return;
    
    let newDate: Date;
    if (timeFilter === 'day') newDate = direction === 'prev' ? subDays(filterBaseDate, 1) : addDays(filterBaseDate, 1);
    else if (timeFilter === 'week') newDate = direction === 'prev' ? subWeeks(filterBaseDate, 1) : addWeeks(filterBaseDate, 1);
    else if (timeFilter === 'month') newDate = direction === 'prev' ? subMonths(filterBaseDate, 1) : addMonths(filterBaseDate, 1);
    else newDate = direction === 'prev' ? subYears(filterBaseDate, 1) : addYears(filterBaseDate, 1);
    
    setFilterBaseDate(newDate);
  };

  // Derived State: Batches (Tracking stock per batch) - OPTIMIZED O(N)
  const batches = useMemo(() => {
    if (!transactions.length) return [];
    
    // Look up products faster with a map
    const productMap = new Map<string, Product>();
    products.forEach(p => productMap.set(p.id, p));

    // Group transactions by product and batch for O(1) lookup
    const batchMap = new Map<string, BatchInfo>();
    
    // Single pass through transactions
    for (const t of transactions) {
      if (!t.batchNumber || !t.productId) continue;
      const key = `${t.productId}_${t.batchNumber}`;
      
      let existing = batchMap.get(key);
      if (!existing) {
        const product = productMap.get(t.productId);
        if (!product) continue;
        
        existing = {
          batchNumber: t.batchNumber,
          productId: t.productId,
          productName: t.productName || product.name,
          category: t.category || product.category,
          stock: 0,
          importDate: t.date
        };
        batchMap.set(key, existing);
      }
      
      const qty = Number(t.quantity) || 0;

      if (t.type === 'IN' || t.type === 'OPENING') {
        existing.stock += qty;
        // Keep earliest import date
        if (t.date && t.date < existing.importDate) {
          existing.importDate = t.date;
        }
      } else if ((t.type === 'OUT' || t.type === 'LOSS' || t.type === 'DAMAGE') && t.status !== 'in_transit') {
        existing.stock -= qty;
        if (t.date && (!existing.lastExportDate || t.date > existing.lastExportDate)) {
          existing.lastExportDate = t.date;
        }
      }
    }

    // Return batches (filter empty if needed, or keep for history)
    return Array.from(batchMap.values()).sort((a, b) => {
      const timeA = new Date(a.importDate).getTime();
      const timeB = new Date(b.importDate).getTime();
      if (isNaN(timeA)) return 1;
      if (isNaN(timeB)) return -1;
      return timeA - timeB;
    });
  }, [transactions, products]);

  // AUTO-FILL FIFO SUGGESTION
  useEffect(() => {
    if (activeTab === 'export' && !loading) {
      const currentItems = newTransaction.items;
      let hasChange = false;
      
      const updatedItems = currentItems.map(item => {
        if (!item.productId) return item;
        const product = products.find(p => p.id === item.productId);
        if (!product) return item;

        // Tìm lô đầu tiên còn hàng
        const oldestBatch = batches.find(b => b.productId === item.productId && b.stock > 0);
        const targetBatch = oldestBatch ? oldestBatch.batchNumber : '';
        
        // Chỉ cập nhật nếu Lô dự kiến khác Lô hiện tại và Lô hiện tại chưa được người dùng nhập thủ công (hoặc rỗng)
        if (targetBatch && item.batchNumber !== targetBatch && (!item.batchNumber || item.batchNumber === '')) {
          hasChange = true;
          return { ...item, batchNumber: targetBatch };
        }
        return item;
      });
      
      if (hasChange) {
        setNewTransaction(prev => ({ ...prev, items: updatedItems }));
      }
    }
  }, [activeTab, batches, loading, products]); // Added products to dependencies

  // Default supplier for import
  useEffect(() => {
    if (activeTab === 'import') {
      const skb = partners.find(p => p.id === 'SKB-BNC' || p.name === 'SKB-BNC');
      if (skb && newTransaction.partnerId !== skb.id) {
        setNewTransaction(prev => ({ ...prev, partnerId: skb.id, type: 'IN' }));
      }
    }
  }, [activeTab, partners, newTransaction.partnerId]);

  const filteredTransactionsForReport = useMemo(() => {
    const bq = batchSearchQuery.toLowerCase().trim();
    if (!bq) return filteredTransactionsByTime;
    return filteredTransactionsByTime.filter(t => t.batchNumber?.toLowerCase().includes(bq));
  }, [filteredTransactionsByTime, batchSearchQuery]);

  // Filter for Gallery
  const galleryMonthOptions = useMemo(() => {
    const dates = transactions
      .filter(t => t.evidencePhotoUrl)
      .map(t => startOfMonth(parseISO(t.date)));
    
    const monthStrings = dates.map(d => format(d, 'yyyy-MM'));
    const uniqueMonths = ([...new Set(monthStrings)] as string[])
      .sort((a, b) => b.localeCompare(a)); // Newest first

    return [
      { value: 'all', label: 'Tất cả các tháng' },
      ...uniqueMonths.map((m: string) => {
        try {
          return {
            value: m,
            label: `Tháng ${format(parse(m, 'yyyy-MM', new Date()), 'MM/yyyy')}`
          };
        } catch (e) {
          return { value: m, label: `Tháng ${m}` };
        }
      })
    ];
  }, [transactions]);

  const filteredGalleryTransactions = useMemo(() => {
    return transactions.filter(t => {
      if (t.type !== galleryFilter || !t.evidencePhotoUrl) return false;
      
      if (galleryMonth !== 'all') {
        try {
          const tDate = parseISO(t.date);
          const monthStr = format(tDate, 'yyyy-MM');
          if (monthStr !== galleryMonth) return false;
        } catch (e) {
          return false;
        }
      }
      
      if (gallerySearchQuery.trim() !== '') {
        const query = gallerySearchQuery.toLowerCase();
        if (galleryFilter === 'IN') {
          // Nhập kho: Tra cứu mã lô
          if (!t.batchNumber?.toLowerCase().includes(query)) {
            return false;
          }
        } else {
          // Xuất kho: Tra cứu đơn vị
          if (!t.partnerName.toLowerCase().includes(query)) {
            return false;
          }
        }
      }
      
      return true;
    });
  }, [transactions, galleryFilter, galleryMonth, gallerySearchQuery]);

  const batchLifecycle = useMemo(() => {
    const bq = batchSearchQuery.toLowerCase().trim();
    if (!bq || bq.length < 2) return null;

    const allMatches = transactions.filter(t => t.batchNumber?.toLowerCase().includes(bq));
    if (allMatches.length === 0) return null;

    const imports = allMatches.filter(t => t.type === 'IN' || t.type === 'OPENING');
    const exports = allMatches.filter(t => t.type === 'OUT');
    
    const totalIn = imports.reduce((sum, t) => sum + t.quantity, 0);
    const totalOut = exports.reduce((sum, t) => sum + t.quantity, 0);
    
    return {
      batchNumber: bq,
      imports,
      exports,
      totalIn,
      totalOut,
      balance: totalIn - totalOut,
      productName: allMatches[0].productName,
      unit: products.find(p => p.id === allMatches[0].productId)?.unit || 'Đơn vị'
    };
  }, [transactions, batchSearchQuery, products]);

  // Derived State: Import/Export Flow Summary
  const flowSummary = useMemo(() => {
    const summaryMap = new Map<string, { 
      id: string; 
      in: number; 
      out: number; 
      inValue: number; 
      outValue: number; 
      productName: string; 
      category: Category; 
      unit: string;
      openingStock: number;
      closingStock: number;
    }>();
    
    products.forEach(p => {
       summaryMap.set(p.id, { 
         id: p.id, 
         in: 0, 
         out: 0, 
         inValue: 0, 
         outValue: 0, 
         productName: p.name, 
         category: p.category, 
         unit: p.unit,
         openingStock: 0,
         closingStock: 0
       });
    });

    // 1. Calculate In/Out for the period
    filteredTransactionsForReport.forEach(t => {
      // Don't count in-transit exports in the "OUT" reports until they are completed
      if (t.type === 'OUT' && t.status === 'in_transit') return;

      const entry = summaryMap.get(t.productId);
      const product = products.find(p => p.id === t.productId);
      if (entry && product) {
        if (t.type === 'IN' || t.type === 'OPENING') {
          entry.in += t.quantity;
          entry.inValue += t.quantity * product.price;
        } else {
          // General OUT categories (OUT, LOSS, DAMAGE)
          entry.out += t.quantity;
          entry.outValue += t.quantity * product.price;
        }
      }
    });

    /* 
    // 1.1 Add Revenue data to Outflows - Removed to avoid double counting
    // Stock flow should follow physical transactions. Revenue is for reconciliation.
    filteredRevenueByTime.forEach(r => {
      const product = products.find(p => p.name === r.productName);
      if (product) {
        const entry = summaryMap.get(product.id);
        if (entry) {
          entry.out += r.quantity;
          entry.outValue += r.totalAmount;
        }
      }
    });
    */

    // 2. Calculate Closing Stock (Stock at the end of period)
    const baseTransactions = batchSearchQuery.trim() 
      ? transactions.filter(t => t.batchNumber?.toLowerCase().includes(batchSearchQuery.toLowerCase().trim()))
      : transactions;

    baseTransactions.forEach(t => {
      // Don't subtract from closing stock if it's still in transit
      if (t.type === 'OUT' && t.status === 'in_transit') return;

      const entry = summaryMap.get(t.productId);
      if (entry) {
        const transDate = parseISO(t.date);
        // If "all" time, closing stock is current stock
        // If filtered, closing stock is stock up to end of period
        const isBeforeOrAtEnd = !dateRange || transDate <= dateRange.end;
        
        if (isBeforeOrAtEnd) {
          const multiplier = (t.type === 'IN' || t.type === 'OPENING') ? 1 : -1;
          entry.closingStock += multiplier * t.quantity;
        }
      }
    });

    // 3. Calculate Opening Stock: Opening = Closing - In + Out
    summaryMap.forEach(entry => {
      entry.openingStock = entry.closingStock - entry.in + entry.out;
    });

    return Array.from(summaryMap.values()).filter(item => {
      if (!batchSearchQuery.trim()) return true;
      return item.in > 0 || item.out > 0 || item.closingStock > 0 || item.openingStock > 0;
    });
  }, [filteredTransactionsForReport, filteredRevenueByTime, transactions, products, dateRange, batchSearchQuery]);

  // Use the time-filtered transactions for inventory/stats
  const inventory = useMemo(() => {
    const invMap = new Map<string, InventoryItem>();
    
    // 1. Initialize for all products
    products.forEach(p => {
      invMap.set(p.id, {
        productId: p.id,
        productName: p.name,
        category: p.category,
        unit: p.unit,
        stock: 0,
        totalLiters: 0,
        minStock: p.minStock
      });
    });

    // 2. Aggregate from pre-calculated batches (O(B))
    const productMap = new Map<string, Product>();
    products.forEach(p => productMap.set(p.id, p));

    batches.forEach(b => {
      const item = invMap.get(b.productId);
      const product = productMap.get(b.productId);
      if (item && product) {
        item.stock += b.stock;
        const units = b.stock * (product.conversionFactor || 1);
        const liters = (units * (product.capacityPerUnit || 1000)) / 1000;
        item.totalLiters += liters;
      }
    });

    return Array.from(invMap.values()).sort((a, b) => b.stock - a.stock);
  }, [batches, products]);

  // Reconciliation Dashboard Logic
  const reconciliationData = useMemo(() => {
    const dataMap = new Map<string, { 
      productId: string; 
      productName: string; 
      exportQty: number; 
      revenueQty: number;
      unit: string;
      category: string;
    }>();

    // 1. Calculate physical exports
    filteredTransactionsByTime.forEach(t => {
      if (t.type === 'OUT') {
        const product = products.find(p => p.id === t.productId);
        const entry = dataMap.get(t.productId) || {
          productId: t.productId,
          productName: t.productName,
          exportQty: 0,
          revenueQty: 0,
          unit: product?.unit || 'ĐV',
          category: product?.category || 'KHÁC'
        };
        entry.exportQty += t.quantity;
        dataMap.set(t.productId, entry);
      }
    });

    // 2. Calculate revenue quantities
    filteredRevenueByTime.forEach(r => {
      const product = products.find(p => p.name === r.productName);
      const productId = product ? product.id : `legacy_${r.productName}`;
      
      const entry = dataMap.get(productId) || {
        productId,
        productName: r.productName,
        exportQty: 0,
        revenueQty: 0,
        unit: r.unit || 'ĐV',
        category: product?.category || 'KHÁC'
      };
      entry.revenueQty += r.quantity;
      dataMap.set(productId, entry);
    });

    return Array.from(dataMap.values()).map(item => ({
      ...item,
      diff: item.exportQty - item.revenueQty,
      matchPercentage: item.exportQty > 0 ? Math.min(100, (item.revenueQty / item.exportQty) * 100) : (item.revenueQty === 0 ? 100 : 0)
    })).sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
  }, [filteredTransactionsByTime, filteredRevenueByTime, products]);

  const totalReconciliationScore = useMemo(() => {
    if (reconciliationData.length === 0) return 100;
    const totalDiff = reconciliationData.reduce((acc, curr) => acc + Math.abs(curr.diff), 0);
    const totalExport = reconciliationData.reduce((acc, curr) => acc + curr.exportQty, 0);
    if (totalExport === 0) return 100;
    return Math.max(0, 100 - (totalDiff / totalExport * 100));
  }, [reconciliationData]);

  // Derived State: Stats & Turnover
  const stats = useMemo(() => {
    const totalIn = filteredTransactionsByTime.filter(t => t.type === 'IN').reduce((acc, curr) => acc + curr.quantity, 0);
    const totalOut = filteredTransactionsByTime.filter(t => 
      (t.type === 'OUT' && (t.status === 'completed' || !t.status)) || 
      t.type === 'LOSS' || 
      t.type === 'DAMAGE'
    ).reduce((acc, curr) => acc + curr.quantity, 0);
    const currentStock = inventory.reduce((acc, curr) => acc + curr.stock, 0);
    const partnerCount = partners.length;
    const totalLiters = inventory.reduce((acc, curr) => acc + curr.totalLiters, 0);

    const avgInventory = totalIn / 2 || 1;
    const turnoverRate = Number((totalOut / avgInventory).toFixed(2));

    const totalValue = inventory.reduce((acc, curr) => {
      const product = products.find(p => p.id === curr.productId);
      return acc + (curr.stock * (product?.price || 0));
    }, 0);

    const lowStockThreshold = 10;
    const lowStockItems = inventory.filter(i => i.stock > 0 && i.stock < lowStockThreshold).length;
    const outOfStockItems = inventory.filter(i => i.stock <= 0).length;
    const healthyItems = inventory.filter(i => i.stock >= lowStockThreshold).length;

    return { 
      totalIn, 
      totalOut, 
      currentStock, 
      partnerCount, 
      totalLiters, 
      turnoverRate, 
      totalValue,
      inventoryHealth: {
        low: lowStockItems,
        out: outOfStockItems,
        healthy: healthyItems
      }
    };
  }, [filteredTransactionsByTime, partners, inventory, products]);

  const filteredTransactions = useMemo(() => {
    const q = historySearchQuery.toLowerCase().trim();
    if (!q) return filteredTransactionsByTime;
    return filteredTransactionsByTime.filter(t => {
      const basicMatch = t.productName.toLowerCase().includes(q) || 
        t.partnerName.toLowerCase().includes(q) || 
        t.batchNumber?.toLowerCase().includes(q) ||
        t.notes?.toLowerCase().includes(q);

      if (basicMatch) return true;

      const partner = partners.find(p => p.id === t.partnerId);
      return partner?.sapCode?.toLowerCase().includes(q);
    });
  }, [filteredTransactionsByTime, historySearchQuery, partners]);

  const chartData = useMemo(() => {
    const categories: Category[] = ['Lon', 'Lít', 'Chai'];
    return categories.map(cat => ({
      name: cat,
      Nhập: filteredTransactionsByTime.filter(t => t.category === cat && t.type === 'IN').reduce((acc, curr) => acc + curr.quantity, 0),
      Xuất: filteredTransactionsByTime.filter(t => t.category === cat && t.type === 'OUT').reduce((acc, curr) => acc + curr.quantity, 0),
      'Doanh thu': filteredRevenueByTime.filter(r => {
        const p = products.find(prod => prod.name === r.productName);
        return p?.category === cat;
      }).reduce((acc, curr) => acc + curr.quantity, 0),
    }));
  }, [filteredTransactionsByTime, filteredRevenueByTime, products]);

  const handleAddTransaction = async (type: TransactionType) => {
    const par = partners.find(partner => partner.id === newTransaction.partnerId);
    
    // Global Validation
    if ((!par && type !== 'OPENING')) {
      alert("Vui lòng chọn đối tác!");
      return;
    }

    // Items Validation
    const validItems = newTransaction.items.filter(item => item.productId && item.quantity > 0);
    if (validItems.length === 0) {
      alert("Vui lòng chọn ít nhất một sản phẩm với số lượng hợp lệ!");
      return;
    }

    if (type === 'IN') {
      const missingBatch = validItems.find(item => !item.batchNumber?.trim());
      if (missingBatch) {
        alert("Tin Tin từ chối: Anh/Chị bắt buộc phải nhập Mã lô cho tất cả mặt hàng khi thực hiện nhập kho ạ!");
        return;
      }
    }

    if (type === 'OUT') {
      for (const item of validItems) {
        const currentItem = inventory.find(i => i.productId === item.productId);
        if (!currentItem || currentItem.stock < item.quantity) {
          alert(`Sản phẩm "${products.find(p => p.id === item.productId)?.name}" không đủ số lượng trong kho!`);
          return;
        }
      }
    }

    setLoading(true);
    try {
      const batch = writeBatch(db);
      const referenceGroupId = `multi-${Date.now()}`;
      const timestamp = new Date().toISOString();
      const transactionDate = newTransaction.date 
        ? (newTransaction.date.includes('T') ? newTransaction.date : `${newTransaction.date}T${timestamp.split('T')[1]}`)
        : timestamp;

      // Local copy of batches to handle multiple items in one go
      let currentBatchesLocal = batches.map(b => ({ ...b }));

      for (const item of validItems) {
        const p = products.find(prod => prod.id === item.productId);
        if (!p) continue;

        if (type === 'OUT') {
          const allocations = getFIFOAllocations(p.id, Number(item.quantity), currentBatchesLocal);
          
          for (let i = 0; i < allocations.length; i++) {
            const alloc = allocations[i];
            const transactionId = `split-${Date.now()}-${item.productId}-${i}`;
            const transaction: Transaction = {
              id: transactionId,
              date: transactionDate,
              type: 'OUT',
              productId: p.id,
              productName: p.name,
              category: p.category,
              quantity: alloc.quantity,
              partnerId: par?.id || 'UNKNOWN',
              partnerName: par?.name || 'Vô danh',
              notes: validItems.length > 1 
                ? `[Món ${validItems.indexOf(item) + 1}/${validItems.length}] ${allocations.length > 1 ? `[Lô ${i+1}/${allocations.length}] ` : ''}${newTransaction.notes}`
                : (allocations.length > 1 ? `[Lô ${i+1}/${allocations.length}] ${newTransaction.notes}` : newTransaction.notes),
              batchNumber: alloc.batchNumber,
              evidencePhotoUrl: newTransaction.evidencePhotoUrl || undefined,
              createdBy: userEmail || user || 'Guest',
              referenceGroupId: referenceGroupId,
              status: newTransaction.isInTransit ? 'in_transit' : 'completed',
              originalQuantity: alloc.quantity
            };
            batch.set(doc(db, 'transactions', transactionId), transaction);
          }
        } else {
          const transactionId = `trx-${Date.now()}-${item.productId}`;
          const transaction: Transaction = {
            id: transactionId,
            date: transactionDate,
            type,
            productId: p.id,
            productName: p.name,
            category: p.category,
            quantity: Number(item.quantity),
            partnerId: type === 'OPENING' ? 'SYSTEM_BEGINNING' : (par?.id || 'UNKNOWN'),
            partnerName: type === 'OPENING' ? 'Số dư đầu kỳ' : (par?.name || 'Vô danh'),
            notes: validItems.length > 1 ? `[Món ${validItems.indexOf(item) + 1}/${validItems.length}] ${newTransaction.notes}` : newTransaction.notes,
            batchNumber: item.batchNumber || undefined,
            evidencePhotoUrl: newTransaction.evidencePhotoUrl || undefined,
            createdBy: userEmail || user || 'Guest',
            referenceGroupId: referenceGroupId
          };
          batch.set(doc(db, 'transactions', transactionId), transaction);
        }
      }

      await batch.commit();
      
      const inTransit = newTransaction.isInTransit;
      
      // Immediate state updates
      setActiveTab(inTransit ? 'in-transit' : 'history');
      setNewTransaction({ 
        type: 'IN',
        partnerId: partners.find(p => p.id === 'SKB-BNC' || p.name === 'SKB-BNC')?.id || partners[0]?.id || '',
        notes: '',
        evidencePhotoUrl: '',
        date: format(new Date(), 'yyyy-MM-dd'),
        isInTransit: false,
        items: [{ productId: products[0]?.id || '', quantity: 0, batchNumber: '' }]
      });
      
      setLoading(false);
      showNotification('Hệ thống cập nhật data thành công');
    } catch (err) {
      console.error(err);
      setLoading(false);
      showNotification('Lỗi khi lưu giao dịch. Anh kiểm tra lại kết nối nhé!', 'error');
    }
  };


  const handleDeleteAllTransactions = async () => {
    if (!isOwner) {
      alert("Chỉ anh Khoa mới có quyền xóa sạch dữ liệu ạ!");
      return;
    }
    
    if (!window.confirm("CẢNH BÁO: Tin Tin sẽ xóa TOÀN BỘ lịch sử giao dịch. Thao tác này KHÔNG THỂ HOÀN TÁC. Anh có chắc chắn không ạ?")) return;
    
    setLoading(true);
    try {
      for (const t of transactions) {
        await deleteDoc(doc(db, 'transactions', t.id));
      }
      setLoading(false);
      showNotification("Đã xóa sạch toàn bộ lịch sử giao dịch.");
    } catch (err) {
      setLoading(false);
      handleFirestoreError(err, OperationType.DELETE, 'transactions');
    }
  };

  const handleDeleteTransaction = async (id: string, productName: string) => {
    if (!isOwner) {
      alert("Chỉ anh Khoa mới có quyền xóa giao dịch ạ!");
      return;
    }
    if (!window.confirm(`Anh có chắc chắn muốn xóa giao dịch của sản phẩm "${productName}" này không?`)) return;
    
    try {
      await deleteDoc(doc(db, 'transactions', id));
      showNotification('Đã xóa giao dịch thành công');
    } catch (err) {
      handleFirestoreError(err, OperationType.DELETE, `transactions/${id}`);
    }
  };

  const isSuperAdmin = useMemo(() => {
    // Everyone signed in can now perform most operations
    return !!user;
  }, [user]);

  const isOwner = useMemo(() => {
    // ONLY the specific owner can delete
    return userEmail?.toUpperCase() === 'KHOA.HUYNH.06.12.2000@GMAIL.COM';
  }, [userEmail]);

  const isAuthorizedFull = useMemo(() => {
    // Everyone signed in can now see all reports and features
    return !!user;
  }, [user]);

  // Nav items configuration
  const navItems = useMemo(() => {
    const allItems = [
      { id: 'dashboard', label: 'Tổng quan', icon: LayoutDashboard, color: '#3b82f6' },
      { id: 'inventory', label: 'Tồn kho', icon: Package, color: '#f59e0b' },
      { id: 'reports', label: 'Báo cáo', icon: TrendingUp, color: '#f43f5e' },
      { id: 'revenue-mgmt', label: 'Doanh thu', icon: FileSpreadsheet, color: '#8b5cf6' },
      { id: 'in-transit', label: 'Đơn đi đường', icon: Truck, color: '#fbbf24' },
      { id: 'import', label: 'Nhập kho', icon: PlusCircle, color: '#10b981' },
      { id: 'export', label: 'Xuất kho', icon: MinusCircle, color: '#f97316' },
      { id: 'gallery', label: 'Thư viện ảnh', icon: ImageIcon, color: '#ec4899' },
      { id: 'partners', label: 'Đối tác', icon: Users, color: '#6366f1' },
      { id: 'history', label: 'Lịch sử', icon: History, color: '#64748b' },
    ];
    
    // Everyone sees reports and revenue management now
    return allItems;
  }, [isAuthorizedFull]);

  // Handle unauthorized tab access
  useEffect(() => {
    const navIds = navItems.map(i => i.id);
    if (user && !navIds.includes(activeTab)) {
      setActiveTab(navIds[0] || 'dashboard');
    }
  }, [user, activeTab, navItems]);

  if (loading && !user) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-6 relative overflow-hidden">
        {/* Superior Background Architecture */}
        <div className="absolute top-0 right-0 w-[1000px] h-[1000px] bg-primary/20 blur-[160px] rounded-full -mr-[500px] -mt-[500px] pointer-events-none" />
        <div className="absolute bottom-0 left-0 w-[800px] h-[800px] bg-emerald-500/10 blur-[130px] rounded-full -ml-[400px] -mb-[400px] pointer-events-none" />
        
        <div className="flex flex-col items-center gap-12 relative z-10">
          <div className="relative">
            <div className="w-24 h-24 border-4 border-white/5 border-t-amber-500 rounded-full animate-spin" />
            <div className="absolute inset-0 flex items-center justify-center">
              <Beer className="w-8 h-8 text-amber-500/50 animate-pulse" />
            </div>
          </div>
          <div className="text-center space-y-4">
            <h1 className="text-4xl font-black text-white tracking-widest font-serif italic uppercase">Bia Bà Nà</h1>
            <p className="text-white/30 text-[10px] font-black uppercase tracking-[0.5em] animate-pulse">Đang kết nối hệ thống dữ liệu...</p>
          </div>
        </div>
      </div>
    );
  }

  if (!user) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-6 relative overflow-hidden font-sans">
        {/* Superior Background Architecture */}
        <div className="absolute top-0 right-0 w-[1000px] h-[1000px] bg-primary/20 blur-[160px] rounded-full -mr-[500px] -mt-[500px] pointer-events-none" />
        <div className="absolute bottom-0 left-0 w-[800px] h-[800px] bg-emerald-500/10 blur-[130px] rounded-full -ml-[400px] -mb-[400px] pointer-events-none" />
        
        <div 
          className="w-full max-w-md relative z-10"
        >
          <div className="bg-white/10 backdrop-blur-3xl rounded-[32px] sm:rounded-[48px] p-8 sm:p-12 shadow-2xl border border-white/10 ring-1 ring-white/5">
            <div className="flex flex-col items-center text-center mb-8 sm:mb-14">
              <div className="w-16 h-16 sm:w-24 sm:h-24 bg-white/5 backdrop-blur-md rounded-[24px] sm:rounded-[36px] flex items-center justify-center border border-white/15 mb-6 sm:mb-10 shadow-inner group relative overflow-hidden">
                <div className="absolute inset-0 bg-gradient-to-br from-amber-400 to-orange-600 opacity-20 group-hover:opacity-40 transition-opacity" />
                <Beer className="w-8 h-8 sm:w-12 sm:h-12 text-amber-400 relative z-10 group-hover:scale-110 transition-transform duration-700" />
              </div>
              <h1 className="text-3xl sm:text-5xl font-black text-white tracking-tight leading-none mb-3 sm:mb-4 font-serif italic selection:bg-white/10 uppercase">Bia Bà Nà</h1>
              <div className="flex items-center gap-2 sm:gap-3">
                <div className="h-[1px] w-3 sm:w-4 bg-white/20" />
                <p className="text-white/40 text-[8px] sm:text-[9px] font-black uppercase tracking-[0.3em] sm:tracking-[0.4em]">SunCraft brewing system</p>
                <div className="h-[1px] w-3 sm:w-4 bg-white/20" />
              </div>
            </div>
            
            <div className="space-y-5 sm:space-y-6">
              <p className="text-center text-white/60 text-[10px] sm:text-[11px] font-medium leading-relaxed mb-4 uppercase tracking-widest">
                Vui lòng đăng nhập để truy cập hệ thống console
              </p>
              
              <div className="flex flex-col gap-3 sm:gap-4">
                <button 
                  onClick={handleLogin}
                  disabled={isAuthenticating}
                  className="w-full py-4 sm:py-5 flex items-center justify-center gap-3 text-[10px] sm:text-[11px] tracking-[0.15em] sm:tracking-[0.2em] uppercase font-black bg-white text-slate-900 border-none hover:bg-slate-100 shadow-2xl transition-all active:scale-[0.98] rounded-xl sm:rounded-2xl disabled:opacity-50 disabled:cursor-wait"
                >
                  {isAuthenticating ? (
                    <RefreshCw className="w-5 h-5 animate-spin" />
                  ) : (
                    <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" className="w-5 h-5" alt="Google" />
                  )}
                  {isAuthenticating ? 'Đang xác thực...' : 'Sign in with Google'}
                </button>

                <div className="flex items-center gap-4 py-2">
                  <div className="h-[1px] flex-1 bg-white/10" />
                  <span className="text-white/20 text-[9px] font-black uppercase tracking-widest">Hoặc</span>
                  <div className="h-[1px] flex-1 bg-white/10" />
                </div>

                {!showGuestNameModal ? (
                  <button 
                    onClick={() => setShowGuestNameModal(true)}
                    className="w-full py-5 flex items-center justify-center gap-3 text-[11px] tracking-[0.2em] uppercase font-black bg-white/5 text-white/80 border border-white/10 hover:bg-white/10 transition-all active:scale-[0.98] rounded-2xl backdrop-blur-md"
                  >
                    <Users className="w-5 h-5 opacity-40" />
                    Tiếp tục với tư cách Khách
                  </button>
                ) : (
                  <div 
                    className="space-y-4 p-6 bg-white/5 rounded-[24px] border border-white/10 backdrop-blur-3xl"
                  >
                    <p className="text-[9px] font-black text-white/40 uppercase tracking-widest mb-2 px-1">Nhập tên của bạn</p>
                    <input 
                      autoFocus
                      placeholder="VD: Nguyễn Văn A..."
                      value={guestNameInput}
                      onChange={(e) => setGuestNameInput(e.target.value)}
                      onKeyDown={(e) => e.key === 'Enter' && handleGuestLogin()}
                      className="w-full px-5 py-4 bg-white/10 border border-white/10 rounded-2xl text-white font-bold outline-none focus:ring-2 focus:ring-primary/50 transition-all placeholder:text-white/20"
                    />
                    <div className="grid grid-cols-2 gap-3">
                      <button 
                        onClick={() => { setShowGuestNameModal(false); setGuestNameInput(''); }}
                        className="py-3.5 text-[10px] font-black uppercase tracking-widest text-white/40 hover:text-white transition-colors"
                      >
                        Hủy
                      </button>
                      <button 
                        onClick={handleGuestLogin}
                        disabled={!guestNameInput.trim()}
                        className="py-3.5 bg-white text-slate-900 rounded-xl text-[10px] font-black uppercase tracking-widest disabled:opacity-20 active:scale-95 transition-all"
                      >
                        Xác nhận
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
            
            <div className="mt-16 pt-10 border-t border-white/5 flex justify-between items-center opacity-10">
              <div className="flex gap-4">
                <div className="w-1.5 h-1.5 rounded-full bg-white" />
                <div className="w-1.5 h-1.5 rounded-full bg-white opacity-50" />
                <div className="w-1.5 h-1.5 rounded-full bg-white opacity-25" />
              </div>
              <span className="text-[9px] font-black text-white uppercase tracking-widest">Cloud Edition v3.0</span>
            </div>
          </div>
          <p className="text-center text-white/10 text-[8px] font-black uppercase tracking-[0.5em] mt-10">Zero-Trust Cloud Security Enforced</p>
        </div>
      </div>
    );
  }

  return (
    <div className={cn(
      "flex h-screen bg-bg-main text-slate-900 font-sans overflow-hidden transition-all duration-300",
      loading && user ? "blur-[2px] pointer-events-none select-none brightness-95" : ""
    )}>
      {/* Processing Overlay */}
      {loading && user && (
        <div className="fixed inset-0 z-[100] flex flex-col items-center justify-center bg-white/5 backdrop-blur-[1px]">
          <div className="bg-white/90 backdrop-blur-3xl p-8 rounded-[40px] shadow-2xl border border-white flex flex-col items-center gap-6 scale-up-center">
            <div className="relative">
              <div className="w-16 h-16 border-[5px] border-slate-100 border-t-primary rounded-full animate-spin" />
              <div className="absolute inset-0 flex items-center justify-center">
                 <RefreshCw className="w-6 h-6 text-primary/30 animate-spin-slow" />
              </div>
            </div>
            <div className="text-center">
              <p className="text-[11px] font-black text-slate-900 uppercase tracking-[0.3em] mb-1">Tin Tin đang xử lý</p>
              <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">Vui lòng đợi trong giây lát...</p>
            </div>
          </div>
        </div>
      )}

      {/* Non-blocking Notification */}
      {notification && (
        <div className="fixed bottom-10 left-1/2 -translate-x-1/2 z-[200] pointer-events-none">
          <div className={cn(
            "px-6 py-3 rounded-2xl shadow-2xl border backdrop-blur-xl animate-in fade-in slide-in-from-bottom-5 duration-300",
            notification.type === 'success' 
              ? "bg-emerald-500/90 text-white border-emerald-400" 
              : "bg-rose-500/90 text-white border-rose-400"
          )}>
            <div className="flex items-center gap-3">
              {notification.type === 'success' ? <CheckCircle className="w-5 h-5" /> : <AlertTriangle className="w-5 h-5" />}
              <span className="font-bold text-sm tracking-tight">{notification.message}</span>
            </div>
          </div>
        </div>
      )}

      {/* Sidebar Overlay (Mobile) */}
      {sidebarOpen && (
        <div 
          onClick={() => setSidebarOpen(false)}
          className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-40 lg:hidden"
        />
      )}

      {/* Sidebar */}
      <aside 
        className={cn(
          "fixed inset-y-0 left-0 z-50 w-64 sm:w-72 bg-white border-r border-slate-100 transition-all duration-300 transform lg:relative lg:translate-x-0 overflow-hidden shrink-0 flex flex-col premium-shadow",
          sidebarOpen ? "translate-x-0" : "-translate-x-full lg:w-0 lg:opacity-0 lg:border-none"
        )}
      >
        <div className="flex flex-col h-full">
          <div className="p-6 sm:p-10">
            <div className="flex items-center gap-3 sm:gap-4">
              <div className="w-10 h-10 sm:w-12 sm:h-12 bg-amber-500 rounded-[14px] sm:rounded-[18px] flex items-center justify-center text-white shadow-xl shadow-amber-500/20 rotate-3 group hover:rotate-0 transition-transform">
                <Beer className="w-5 h-5 sm:w-6 sm:h-6" />
              </div>
              <div className="min-w-0">
                <h1 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight font-serif italic leading-none">BIA BÀ NÀ</h1>
                <p className="text-[8px] sm:text-[9px] font-black text-slate-400 uppercase tracking-[0.2em] mt-1 sm:mt-2">SUNCRAFT BREWERY</p>
              </div>
            </div>
          </div>

          <nav className="flex-1 px-3 sm:px-5 space-y-2 sm:space-y-3 custom-scrollbar overflow-y-auto pt-2 sm:pt-4">
            <p className="px-4 sm:px-5 text-[8px] sm:text-[10px] font-black text-slate-300 uppercase tracking-[0.2em] sm:tracking-[0.25em] mb-3 sm:mb-4">Command Center</p>
            {navItems.map((item) => (
              <button
                key={item.id}
                onClick={() => {
                  setActiveTab(item.id);
                  if (window.innerWidth < 1024) setSidebarOpen(false);
                }}
                className={cn(
                  "w-full flex items-center gap-3 sm:gap-4 px-4 sm:px-5 py-3.5 sm:py-4.5 rounded-xl sm:rounded-2xl transition-all duration-300 group relative overflow-hidden",
                  activeTab === item.id 
                    ? "bg-slate-900 text-white shadow-2xl shadow-slate-200" 
                    : "text-slate-500 hover:bg-slate-50 hover:text-slate-900"
                )}
              >
                <item.icon 
                  className={cn("w-5 h-5 sm:w-6 sm:h-6 transition-transform duration-300 group-hover:scale-110")} 
                  style={{ color: activeTab === item.id ? '#ffffff' : item.color }}
                />
                <span className="text-xs sm:text-sm font-black uppercase tracking-widest">{item.label}</span>
                {activeTab === item.id && (
                  <div 
                    className="absolute inset-0 bg-gradient-to-r"
                    style={{ background: `linear-gradient(to right, ${item.color}20, transparent)` }}
                  />
                )}
              </button>
            ))}
          </nav>

          <div className="px-6 py-8">
            <div className="bg-slate-50/50 backdrop-blur-sm rounded-[24px] p-6 border border-slate-100/50 flex flex-col gap-5">
              <div className="flex items-center gap-4">
                <div className="w-12 h-12 rounded-[16px] bg-white flex items-center justify-center text-primary font-black shadow-sm ring-1 ring-slate-100 border-b-2 border-slate-50">
                  {user?.charAt(0).toUpperCase()}
                </div>
                <div className="flex-1 min-w-0">
                  <p className="text-xs font-black text-slate-900 truncate tracking-tight">{user}</p>
                  <div className="flex items-center gap-1.5 mt-1">
                    <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                    <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest truncate">{user === 'Guest' ? 'Guest Access' : (user === 'ADMIN' ? 'Full Authority' : 'Expert Analyst')}</p>
                  </div>
                </div>
              </div>
              <div className="h-[1px] bg-slate-100" />
              {isOwner && (
                <button 
                  onClick={handleHardReset}
                  className="w-full flex items-center justify-center gap-3 py-3 rounded-xl text-amber-600 hover:bg-amber-50 text-[10px] font-black uppercase tracking-[0.2em] transition-all border border-dashed border-amber-200 hover:border-amber-300 mb-2"
                >
                  <RefreshCw className="w-4 h-4 animate-spin-slow" /> Dọn sạch hệ thống
                </button>
              )}
              <button 
                onClick={handleLogout}
                className="w-full flex items-center justify-center gap-3 py-3 rounded-xl text-rose-500 hover:bg-rose-50 text-[10px] font-black uppercase tracking-[0.2em] transition-all border border-transparent hover:border-rose-100"
              >
                <LogOut className="w-4 h-4" /> Terminate Session
              </button>
            </div>
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col min-w-0 overflow-hidden relative">
        {/* Header */}
        <header className="h-16 sm:h-[72px] bg-white/80 backdrop-blur-md border-b border-slate-100 flex items-center justify-between px-4 sm:px-8 sticky top-0 z-30 shrink-0">
          <div className="flex items-center gap-3 sm:gap-4">
            <button 
              onClick={() => setSidebarOpen(!sidebarOpen)}
              className="p-2 sm:p-2.5 rounded-lg sm:rounded-xl hover:bg-slate-50 transition-colors"
            >
              <Menu className="w-5 h-5 text-slate-600" />
            </button>
            <div className="flex items-center gap-2">
               <h2 className="text-base sm:text-lg font-extrabold text-slate-900 tracking-tight">
                {navItems.find(i => i.id === activeTab)?.label}
              </h2>
            </div>
          </div>

          <div className="flex items-center gap-2 sm:gap-6">

            <div className="hidden md:flex items-center gap-3 px-4 py-2 bg-slate-50 rounded-xl border border-slate-100">
              <Calendar className="w-4 h-4 text-primary" />
              <span className="text-xs font-bold text-slate-600">{new Date().toLocaleDateString('vi-VN', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</span>
            </div>
          </div>
        </header>

        {/* Scrollable Content */}
        <div className="flex-1 overflow-y-auto p-4 sm:p-8 custom-scrollbar bg-slate-50/30">
          <div
            key={activeTab}
            className="max-w-7xl mx-auto space-y-4 sm:space-y-8 pb-24"
          >
              {/* Global Filter Bar for Analytical Tabs */}
              {['dashboard', 'inventory', 'reports', 'revenue-mgmt', 'history'].includes(activeTab) && (
                <div className="bg-white/80 backdrop-blur-md p-3 sm:p-4 rounded-[22px] sm:rounded-[28px] border border-slate-100 shadow-xl shadow-slate-200/50 flex flex-col md:flex-row items-center justify-between gap-4">
                  <div className="flex bg-slate-100/50 p-1 rounded-xl sm:rounded-2xl border border-slate-100 w-full md:w-auto overflow-x-auto no-scrollbar">
                    {[
                      { id: 'all', label: 'Tất cả' },
                      { id: 'day', label: 'Ngày thực tế' },
                      { id: 'week', label: 'Tuần thực tế' },
                      { id: 'month', label: 'Tháng thực tế' },
                      { id: 'year', label: 'Năm thực tế' },
                    ].map(f => (
                      <button
                        key={f.id}
                        onClick={() => setTimeFilter(f.id as any)}
                        className={cn(
                          "flex-1 md:flex-none px-3 sm:px-6 py-2 sm:py-2.5 rounded-lg sm:rounded-xl text-[9px] sm:text-[10px] font-black uppercase tracking-widest transition-all whitespace-nowrap",
                          timeFilter === f.id 
                            ? "bg-white text-primary shadow-sm ring-1 ring-slate-200" 
                            : "text-slate-400 hover:text-slate-600"
                        )}
                      >
                        {f.label}
                      </button>
                    ))}
                  </div>

                  {timeFilter !== 'all' && (
                    <div className="flex items-center gap-4 bg-slate-50/50 px-6 py-2 rounded-2xl border border-slate-100">
                      <button 
                        onClick={() => moveFilterDate('prev')}
                        className="p-1.5 rounded-lg hover:bg-white hover:shadow-sm text-slate-400 hover:text-primary transition-all"
                      >
                        <ChevronRight className="w-5 h-5 rotate-180" />
                      </button>
                      
                      <div className="flex flex-col items-center min-w-[140px] relative group px-2">
                        <button 
                          onClick={() => {
                            try {
                              if (dateInputRef.current) {
                                if ('showPicker' in dateInputRef.current) {
                                  dateInputRef.current.showPicker();
                                } else {
                                  dateInputRef.current.click();
                                }
                              }
                            } catch (err) {
                              console.error('Date picker error:', err);
                              dateInputRef.current?.click();
                            }
                          }}
                          className="flex flex-col items-center hover:scale-105 transition-transform"
                        >
                          <div className="flex items-center gap-1.5">
                            <Calendar className="w-3 h-3 text-primary opacity-60" />
                            <span className="text-xs font-black text-slate-900 uppercase tracking-widest">{periodLabel}</span>
                          </div>
                        </button>
                        <input 
                          ref={dateInputRef}
                          type="date"
                          className="absolute inset-0 opacity-0 pointer-events-none"
                          value={format(filterBaseDate, 'yyyy-MM-dd')}
                          onChange={(e) => {
                            if (e.target.value) setFilterBaseDate(parseISO(e.target.value));
                          }}
                        />
                        <button 
                          onClick={() => setFilterBaseDate(new Date())}
                          className="text-[9px] font-bold text-primary hover:underline mt-0.5"
                        >
                          Về hôm nay
                        </button>
                      </div>

                      <button 
                        onClick={() => moveFilterDate('next')}
                        className="p-1.5 rounded-lg hover:bg-white hover:shadow-sm text-slate-400 hover:text-primary transition-all"
                      >
                        <ChevronRight className="w-5 h-5" />
                      </button>
                    </div>
                  )}

                  <div className="flex items-center gap-3">
                    {isAuthorizedFull && activeTab === 'dashboard' && (
                      <div className="flex gap-2">
                        <label className="flex items-center gap-2 px-4 py-2 bg-primary/10 text-primary rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-primary/20 transition-all cursor-pointer border border-primary/20">
                          <FileUp className="w-3.5 h-3.5" /> Nạp Excel Tồn Kho
                          <input type="file" accept=".xlsx,.xls" onChange={handleImportInventoryExcel} className="hidden" />
                        </label>
                      </div>
                    )}

                    <div className="text-right hidden sm:block">
                      <p className="text-[9px] font-extrabold text-slate-400 uppercase tracking-widest leading-none">Phạm vi dữ liệu</p>
                      <p className="text-[11px] font-black text-slate-900 mt-1 uppercase">Đã tối ưu hóa</p>
                    </div>
                    <div className="w-10 h-10 bg-emerald-50 rounded-xl flex items-center justify-center text-emerald-600 border border-emerald-100">
                      <Filter className="w-4 h-4" />
                    </div>
                  </div>
                </div>
              )}

              {activeTab === 'dashboard' && (
                <>
                  {isAuthorizedFull && (
                    <div className="grid grid-cols-2 md:grid-cols-2 lg:grid-cols-5 gap-3 sm:gap-5">
                    <StatCard 
                      title="TỔNG DOANH THU" 
                      value={formatNumber(filteredRevenueByTime.reduce((a, b) => a + b.totalAmount, 0))} 
                      unit="đ"
                      icon={DollarSign}
                      color="green"
                      subtitle="Dữ liệu theo kỳ"
                      trend={filteredRevenueByTime.length > 0 ? `+${filteredRevenueByTime.length} GD` : "0 GD"}
                      chartData={filteredRevenueByTime.slice(-10).map(r => ({ value: r.totalAmount }))}
                    />
                    <StatCard 
                      title="VÒNG QUAY TỒN KHO" 
                      value={stats.turnoverRate} 
                      unit="x"
                      icon={RefreshCw}
                      color="primary"
                      target="4.5x"
                      trend="+0.3"
                      chartData={[
                        { value: 4.0 }, { value: 4.1 }, { value: 3.9 }, { value: 4.2 }, { value: 4.2 }
                      ]}
                    />
                    <StatCard 
                      title="TỶ LỆ LẤY ĐẦY (FILL RATE)" 
                      value="98.5" 
                      unit="%"
                      icon={Layers}
                      color="green"
                      target="98%"
                      trend="+1.2%"
                      chartData={[
                        { value: 95 }, { value: 96 }, { value: 97.5 }, { value: 98 }, { value: 98.5 }
                      ]}
                    />
                    <StatCard 
                      title="ĐỘ CHÍNH XÁC TỒN KHO" 
                      value="99.8" 
                      unit="%"
                      icon={ShieldCheck}
                      color="primary"
                      target="100%"
                      trend="Ổn định"
                      chartData={[
                        { value: 99.5 }, { value: 99.7 }, { value: 99.8 }, { value: 99.6 }, { value: 99.8 }
                      ]}
                    />
                    <StatCard 
                      title="CHI PHÍ LƯU KHO" 
                      value="12.4" 
                      unit="%"
                      icon={Package2}
                      color="rose"
                      target="Giảm 2% so với Q1"
                      trend="-2.1%"
                      chartData={[
                        { value: 14.5 }, { value: 14.0 }, { value: 13.2 }, { value: 12.8 }, { value: 12.4 }
                      ]}
                    />
                  </div>
                )}

                  <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
                    {/* Operations & Storage Center */}
                    <Card title="Vận hành Kho & Bảo quản" className="lg:col-span-8">
                       <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 sm:gap-6">
                          <div className="p-4 sm:p-6 bg-slate-50/50 rounded-2xl border border-dotted border-slate-200 group hover:bg-white hover:shadow-xl hover:shadow-slate-200/50 transition-all duration-500">
                             <p className="text-[9px] sm:text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2">Nhiệt độ kho lạnh</p>
                             <div className="flex items-baseline gap-2">
                                <h5 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">4.2°C</h5>
                                <span className="text-[9px] sm:text-[10px] font-bold text-slate-400 uppercase">Celsius</span>
                             </div>
                             <div className="mt-3 sm:mt-4 flex items-center gap-2">
                                <div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                                <span className="text-[9px] sm:text-[10px] font-black text-emerald-600 uppercase tracking-widest leading-none">OPTIMAL</span>
                             </div>
                          </div>

                          <div className="p-4 sm:p-6 bg-slate-50/50 rounded-2xl border border-dotted border-slate-200 group hover:bg-white hover:shadow-xl hover:shadow-slate-200/50 transition-all duration-500">
                             <p className="text-[9px] sm:text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2">Độ ẩm trung bình</p>
                             <div className="flex items-baseline gap-2">
                                <h5 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">65%</h5>
                                <span className="text-[9px] sm:text-[10px] font-bold text-slate-400 uppercase">Humidity</span>
                             </div>
                             <div className="mt-3 sm:mt-4 flex items-center gap-2">
                                <div className="w-2 h-2 rounded-full bg-emerald-500" />
                                <span className="text-[9px] sm:text-[10px] font-black text-emerald-600 uppercase tracking-widest leading-none">NORMAL</span>
                             </div>
                          </div>

                          <div className="p-4 sm:p-6 bg-slate-50/50 rounded-2xl border border-dotted border-slate-200 group hover:bg-white hover:shadow-xl hover:shadow-slate-200/50 transition-all duration-500">
                             <p className="text-[9px] sm:text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2">Sử dụng diện tích</p>
                             <div className="flex items-baseline gap-2">
                                <h5 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">78.5%</h5>
                                <span className="text-[9px] sm:text-[10px] font-bold text-slate-400 uppercase">Capacity</span>
                             </div>
                             <div className="mt-3 sm:mt-4 flex items-center gap-2">
                                <div className="w-2 h-2 rounded-full bg-emerald-500" />
                                <span className="text-[9px] sm:text-[10px] font-black text-emerald-600 uppercase tracking-widest leading-none">STABLE</span>
                             </div>
                          </div>

                          <div className="p-4 sm:p-6 bg-slate-50/50 rounded-2xl border border-dotted border-slate-200 group hover:bg-white hover:shadow-xl hover:shadow-slate-200/50 transition-all duration-500">
                             <p className="text-[9px] sm:text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2">Lô hàng chậm (Aging &gt; 15 ngày)</p>
                             <div className="flex items-baseline gap-2">
                                <h5 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">
                                  {batches.filter(b => b.stock > 0 && differenceInDays(new Date(), parseISO(b.importDate)) > 15).length}
                                </h5>
                                <span className="text-[9px] sm:text-[10px] font-bold text-slate-400 uppercase">Lô hàng</span>
                             </div>
                             <div className="mt-3 sm:mt-4 flex items-center gap-2">
                                <div className={cn(
                                  "w-2 h-2 rounded-full",
                                  batches.filter(b => b.stock > 0 && differenceInDays(new Date(), parseISO(b.importDate)) > 15).length > 0 ? "bg-rose-500 animate-pulse" : "bg-emerald-500"
                                )} />
                                <span className={cn(
                                  "text-[9px] sm:text-[10px] font-black uppercase tracking-widest leading-none",
                                  batches.filter(b => b.stock > 0 && differenceInDays(new Date(), parseISO(b.importDate)) > 15).length > 0 ? "text-rose-600" : "text-emerald-600"
                                )}>
                                  {batches.filter(b => b.stock > 0 && differenceInDays(new Date(), parseISO(b.importDate)) > 15).length > 0 ? 'CRITICAL' : 'ALL CLEAR'}
                                </span>
                             </div>
                          </div>
                       </div>
                    </Card>

                    {/* Stock-out & Low Inventory Defense Center */}
                    <Card title="⚠️ Cảnh báo Đứt hàng & Tồn thấp" className="lg:col-span-4 bg-slate-900 border-none shadow-2xl">
                      <div className="space-y-4 sm:space-y-6">
                         {inventory.filter(i => i.stock <= i.minStock).length === 0 ? (
                           <div className="py-8 sm:py-12 flex flex-col items-center text-center">
                              <ShieldCheck className="w-12 h-12 sm:w-16 sm:h-16 text-emerald-500 mb-4 sm:mb-6 opacity-20" />
                              <p className="text-[10px] sm:text-sm font-bold text-white uppercase tracking-widest">Toàn bộ kho hàng ổn định</p>
                           </div>
                         ) : (
                           inventory.filter(i => i.stock <= i.minStock).slice(0, 5).map(item => (
                             <div key={item.productId} className="group cursor-help">
                                <div className="flex justify-between items-center mb-2 sm:mb-3">
                                   <div className="flex flex-col">
                                      <span className="text-xs sm:text-sm font-black text-white uppercase tracking-tight group-hover:text-amber-400 transition-colors leading-none">{item.productName}</span>
                                      <span className="text-[9px] sm:text-[10px] font-bold text-slate-500 uppercase mt-1 sm:mt-1.5">Min: {item.minStock} {item.unit}</span>
                                   </div>
                                   <div className="text-right">
                                      <span className={cn(
                                        "text-sm sm:text-base font-black font-mono",
                                        item.stock <= 0 ? "text-rose-500" : "text-amber-500"
                                      )}>
                                        {item.stock} <span className="text-[9px] sm:text-[10px] opacity-40">{item.unit}</span>
                                      </span>
                                   </div>
                                </div>
                                <div className="h-1.5 sm:h-2 w-full bg-white/5 rounded-full overflow-hidden">
                                   <div 
                                      className={cn(
                                        "h-full rounded-full transition-all duration-1000 ease-out",
                                        item.stock <= 0 ? "bg-rose-500 shadow-[0_0_10px_rgba(244,63,94,0.5)]" : "bg-amber-500 shadow-[0_0_10px_rgba(245,158,11,0.5)]"
                                      )}
                                      style={{ width: `${Math.min(100, (item.stock / (item.minStock || 1)) * 100)}%` }}
                                   />
                                </div>
                             </div>
                           ))
                         )}
                         
                         {inventory.filter(i => i.stock <= i.minStock).length > 5 && (
                           <button 
                             onClick={() => setActiveTab('inventory')}
                             className="w-full py-3 sm:py-4 bg-white/5 hover:bg-white/10 rounded-xl sm:rounded-2xl text-[10px] sm:text-xs font-black text-slate-400 uppercase tracking-[0.2em] transition-all border border-white/5"
                           >
                             Tất cả {inventory.filter(i => i.stock <= i.minStock).length} cảnh báo
                           </button>
                         )}
                      </div>
                    </Card>

                    {/* Batch Aging & Slow Moving Inventory Dashboard */}
                    <Card title="📊 Phân tích Lô hàng mới (Trong 7 ngày)" className="lg:col-span-12">
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
                        {/* Aging Categories Summary */}
                        <div className="bg-emerald-50/50 p-4 rounded-2xl border border-emerald-100 flex flex-col items-center text-center">
                          <p className="text-[9px] font-black text-emerald-600 uppercase tracking-widest mb-1">Tuân thủ FIFO (Mới)</p>
                          <p className="text-2xl font-black text-emerald-700">{batches.filter(b => b.stock > 0 && differenceInDays(new Date(), parseISO(b.importDate)) < 7).length}</p>
                          <p className="text-[8px] font-bold text-emerald-500 mt-1 uppercase">Dưới 7 ngày</p>
                        </div>
                        <div className="bg-amber-50/50 p-4 rounded-2xl border border-amber-100 flex flex-col items-center text-center">
                          <p className="text-[9px] font-black text-amber-600 uppercase tracking-widest mb-1">Cảnh báo chậm (Lưu kho)</p>
                          <p className="text-2xl font-black text-amber-700">{batches.filter(b => b.stock > 0 && differenceInDays(new Date(), parseISO(b.importDate)) >= 7 && differenceInDays(new Date(), parseISO(b.importDate)) <= 15).length}</p>
                          <p className="text-[8px] font-bold text-amber-500 mt-1 uppercase">7 - 15 ngày</p>
                        </div>
                        <div className="bg-rose-50/50 p-4 rounded-2xl border border-rose-100 flex flex-col items-center text-center">
                          <p className="text-[9px] font-black text-rose-600 uppercase tracking-widest mb-1">Hàng tồn đọng (Critical)</p>
                          <p className="text-2xl font-black text-rose-700">{batches.filter(b => b.stock > 0 && differenceInDays(new Date(), parseISO(b.importDate)) > 15).length}</p>
                          <p className="text-[8px] font-bold text-rose-500 mt-1 uppercase">Trên 15 ngày</p>
                        </div>
                      </div>

                      <div className="overflow-x-auto">
                        <table className="w-full text-left">
                          <thead>
                            <tr className="bg-slate-50/20">
                              <th className="font-serif italic text-slate-400 text-[10px] sm:text-xs uppercase tracking-widest py-3 sm:py-5 px-3 sm:px-6">Mã Lô</th>
                              <th className="font-serif italic text-slate-400 text-[10px] sm:text-xs uppercase tracking-widest py-3 sm:py-5 px-3 sm:px-6">Sản phẩm</th>
                              <th className="font-serif italic text-slate-400 text-[10px] sm:text-xs uppercase tracking-widest py-3 sm:py-5 px-3 sm:px-6 text-right">Tồn hiện tại</th>
                              <th className="font-serif italic text-slate-400 text-[10px] sm:text-xs uppercase tracking-widest py-3 sm:py-5 px-3 sm:px-6 text-center">Ngày nhập</th>
                              <th className="font-serif italic text-slate-400 text-[10px] sm:text-xs uppercase tracking-widest py-3 sm:py-5 px-3 sm:px-6 text-right">Thời gian lưu kho</th>
                              <th className="font-serif italic text-slate-400 text-[10px] sm:text-xs uppercase tracking-widest py-3 sm:py-5 px-3 sm:px-6 text-right">Trạng thái</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100/50">
                            {batches
                              .filter(b => b.stock > 0 && differenceInDays(new Date(), parseISO(b.importDate)) <= 7)
                              .sort((a, b) => new Date(b.importDate).getTime() - new Date(a.importDate).getTime()) // Newest first for recent view
                              .slice(0, 7)
                              .map((batch) => {
                                const age = differenceInDays(new Date(), parseISO(batch.importDate));
                                return (
                                  <tr key={`${batch.productId}_${batch.batchNumber}`} className="hover:bg-slate-50/50 transition-all group">
                                    <td className="py-3 sm:py-5 px-3 sm:px-6">
                                      <div className="flex items-center gap-3">
                                        <div className={cn(
                                          "w-2 h-8 rounded-full shadow-sm",
                                          age > 15 ? "bg-rose-500 shadow-rose-200" : age >= 7 ? "bg-amber-400 shadow-amber-100" : "bg-emerald-400 shadow-emerald-100"
                                        )} />
                                        <span className="font-mono text-[10px] sm:text-xs text-slate-600 font-black">#{batch.batchNumber}</span>
                                      </div>
                                    </td>
                                    <td className="py-3 sm:py-5 px-3 sm:px-6">
                                      <div className="flex flex-col">
                                        <span className="text-xs sm:text-sm font-black text-slate-900 leading-none tracking-tight">{batch.productName}</span>
                                        <span className="text-[8px] sm:text-[10px] font-bold text-slate-400 uppercase tracking-widest mt-1 sm:mt-1.5">{batch.category}</span>
                                      </div>
                                    </td>
                                    <td className="py-3 sm:py-5 px-3 sm:px-6 font-mono text-right text-xs sm:text-sm font-black text-slate-900">
                                      {formatNumber(batch.stock)}
                                    </td>
                                    <td className="py-3 sm:py-5 px-3 sm:px-6 text-center font-mono text-[9px] sm:text-xs text-slate-400 font-bold">
                                      {formatDisplayDate(batch.importDate)}
                                    </td>
                                    <td className="py-3 sm:py-5 px-3 sm:px-6 text-right">
                                      <span className={cn(
                                        "text-xs sm:text-sm font-black italic",
                                        age > 15 ? "text-rose-600" : age >= 7 ? "text-amber-600" : "text-emerald-600"
                                      )}>
                                        {age} <span className="text-[9px] not-italic opacity-60">ngày</span>
                                      </span>
                                    </td>
                                    <td className="py-3 sm:py-5 px-3 sm:px-6 text-right">
                                      <span className={cn(
                                        "px-2 py-0.5 sm:px-3 sm:py-1 rounded-full text-[8px] sm:text-[9px] font-black uppercase tracking-widest border inline-flex items-center gap-1 shadow-sm",
                                        age > 15 ? "bg-rose-50 border-rose-100 text-rose-600" : 
                                        age >= 7 ? "bg-amber-50 border-amber-100 text-amber-600" : 
                                        "bg-emerald-50 border-emerald-100 text-emerald-600"
                                      )}>
                                        {age > 15 ? 'Đọng kho' : age >= 7 ? 'Chậm' : 'Bình thường'}
                                      </span>
                                    </td>
                                  </tr>
                                );
                              })}
                            {batches.filter(b => b.stock > 0).length === 0 && (
                              <tr>
                                <td colSpan={6} className="py-20 text-center">
                                  <AlertCircle className="w-10 h-10 text-slate-100 mx-auto mb-4" />
                                  <p className="text-xs font-bold text-slate-300 uppercase tracking-widest">Không có lô hàng tồn kho hiện tại</p>
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                      
                      <div className="mt-8 p-5 bg-slate-50/50 rounded-2xl border border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-4">
                        <div className="flex items-center gap-3">
                          <Info className="w-5 h-5 text-primary" />
                          <div className="space-y-1">
                            <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest leading-loose">
                              Dashboard hiện đang lọc các lô hàng nhập trong vòng <span className="text-primary font-black">7 ngày qua</span> để tối ưu hiển thị.
                            </p>
                            {batches.filter(b => b.stock > 0 && differenceInDays(new Date(), parseISO(b.importDate)) > 7).length > 0 && (
                              <p className="text-[9px] font-bold text-rose-500 uppercase tracking-widest">
                                Chú ý: Có {batches.filter(b => b.stock > 0 && differenceInDays(new Date(), parseISO(b.importDate)) > 7).length} lô hàng khác đang có dấu hiệu chậm/đọng kho (&gt; 7 ngày).
                              </p>
                            )}
                          </div>
                        </div>
                        <Button 
                          variant="outline" 
                          onClick={() => setActiveTab('inventory')}
                          className="bg-white text-[10px] w-full sm:w-auto"
                        >
                          KIỂM TRA TOÀN KHO
                        </Button>
                      </div>
                    </Card>
                  </div>
                </>
              )}

              {activeTab === 'inventory' && (
                <div className="space-y-6">
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                    <div className="relative max-w-md w-full group">
                      <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 group-focus-within:text-primary transition-colors hover:text-primary" />
                      <input 
                        placeholder="Tìm kiếm sản phẩm trong kho..." 
                        className="w-full pl-11 pr-4 py-3 bg-white border border-slate-100 rounded-2xl focus:ring-4 focus:ring-primary/5 focus:border-primary outline-none text-sm transition-all premium-shadow"
                      />
                    </div>
                    <Button variant="outline" className="bg-white">
                      <Download className="w-4 h-4" /> Xuất báo cáo Excel
                    </Button>
                  </div>

                  <Card title="Trạng thái tồn kho chi tiết" noPadding>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left">
                        <thead>
                          <tr className="bg-slate-50/50">
                             <th className="font-bold text-[9px] sm:text-[10px] text-slate-400 uppercase tracking-widest py-3 sm:py-4 px-3 sm:px-6">Sản phẩm</th>
                             <th className="font-bold text-[9px] sm:text-[10px] text-slate-400 uppercase tracking-widest py-3 sm:py-4 px-3 sm:px-6">Quy cách</th>
                             <th className="font-bold text-[9px] sm:text-[10px] text-slate-400 uppercase tracking-widest py-3 sm:py-4 px-3 sm:px-6 text-right">Tồn kho</th>
                             <th className="font-bold text-[9px] sm:text-[10px] text-slate-400 uppercase tracking-widest py-3 sm:py-4 px-3 sm:px-6 text-right">Phòng kho</th>
                             <th className="font-bold text-[9px] sm:text-[10px] text-slate-400 uppercase tracking-widest py-3 sm:py-4 px-3 sm:px-6 text-right">Cảnh báo</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {inventory.map((item) => (
                            <tr 
                              key={item.productId} 
                              onClick={() => setSelectedInventoryProduct(item.productId)}
                              className="hover:bg-slate-50 transition-colors group cursor-pointer"
                            >
                              <td className="py-3 sm:py-4 px-3 sm:px-6">
                                <p className="font-bold text-slate-900 leading-none text-xs sm:text-sm">{item.productName}</p>
                                <p className="text-[9px] font-mono font-bold text-slate-400 mt-1 uppercase tracking-widest">{item.productId}</p>
                              </td>
                              <td className="py-3 sm:py-4 px-3 sm:px-6">
                                <span className={cn(
                                  "px-2 py-0.5 sm:px-2.5 sm:py-1 rounded-lg text-[8px] sm:text-[10px] font-black uppercase tracking-tighter border",
                                  item.category === 'Lon' ? "bg-orange-50 border-orange-100 text-orange-600" :
                                  item.category === 'Lít' ? "bg-blue-50 border-blue-100 text-blue-600" : "bg-emerald-50 border-emerald-100 text-emerald-600"
                                )}>
                                  {item.category}
                                </span>
                              </td>
                              <td className="py-3 sm:py-4 px-3 sm:px-6 font-mono text-right text-sm sm:text-lg font-black text-slate-900">
                                {formatNumber(item.stock)}
                              </td>
                              <td className="py-3 sm:py-4 px-3 sm:px-6 font-mono text-right text-sm sm:text-lg font-black text-primary">
                                {formatNumber(item.totalLiters)}<span className="text-[10px] sm:text-xs font-bold ml-0.5">L</span>
                              </td>
                              <td className="py-4 px-6 text-right">
                                {item.stock <= 50 ? (
                                  <span className="bg-rose-50 text-rose-600 border border-rose-100 px-3 py-1 rounded-lg text-[10px] font-black uppercase tracking-widest animate-pulse">
                                    MỨC THẤP
                                  </span>
                                ) : (
                                  <span className="bg-emerald-50 text-emerald-600 border border-emerald-100 px-3 py-1 rounded-lg text-[10px] font-black uppercase tracking-widest">ỔN ĐỊNH</span>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </Card>

                  {/* Batch Detail Modal/Section */}
                  {selectedInventoryProduct && (
                    <div 
                      className="fixed inset-0 z-[100] flex items-center justify-center p-4 sm:p-6"
                    >
                      <div 
                        className="absolute inset-0 bg-slate-900/60 backdrop-blur-md" 
                        onClick={() => setSelectedInventoryProduct(null)} 
                      />
                      <Card className="relative w-full max-w-2xl bg-white shadow-2xl border-2 border-slate-900 rounded-[32px] overflow-hidden" noPadding>
                          <div className="bg-slate-900 p-6 flex justify-between items-center">
                            <div className="flex items-center gap-4">
                              <div className="w-10 h-10 rounded-xl bg-amber-400 flex items-center justify-center">
                                <Package className="w-5 h-5 text-slate-900" />
                              </div>
                              <div>
                                <h3 className="text-white text-lg font-black uppercase tracking-tight italic font-serif"> CHI TIẾT TỒN KHO THEO LÔ</h3>
                                <p className="text-amber-400 text-[9px] font-black uppercase tracking-widest">{products.find(p => p.id === selectedInventoryProduct)?.name}</p>
                              </div>
                            </div>
                            <button 
                              onClick={() => setSelectedInventoryProduct(null)}
                              className="w-10 h-10 rounded-full bg-white/10 hover:bg-white/20 flex items-center justify-center text-white transition-all"
                            >
                              <RefreshCw className="w-5 h-5 rotate-45" />
                            </button>
                          </div>
                          
                          <div className="max-h-[60vh] overflow-y-auto p-4 sm:p-8 custom-scrollbar">
                            <div className="space-y-4">
                              {batches.filter(b => b.productId === selectedInventoryProduct && b.stock > 0).length > 0 ? (
                                batches.filter(b => b.productId === selectedInventoryProduct && b.stock > 0).map((batch) => (
                                  <div 
                                    key={batch.batchNumber} 
                                    className="p-5 rounded-2xl border border-slate-100 bg-slate-50/50 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 group hover:border-primary/30 transition-all"
                                  >
                                    <div className="flex items-center gap-4">
                                      <div className="w-12 h-12 rounded-xl bg-white border border-slate-100 flex flex-col items-center justify-center shadow-sm">
                                        <span className="text-[10px] font-black text-slate-400 uppercase leading-none mb-1">LÔ</span>
                                        <span className="text-xs font-black text-slate-900 font-mono italic">#{batch.batchNumber}</span>
                                      </div>
                                      <div>
                                        <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest leading-none mb-1.5">Ngày nhập</p>
                                        <p className="text-sm font-bold text-slate-900 font-mono">{formatDate(batch.importDate)}</p>
                                      </div>
                                    </div>
                                    
                                    <div className="text-right flex flex-row sm:flex-col items-center sm:items-end justify-between w-full sm:w-auto">
                                      <div className="sm:text-right">
                                        <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest leading-none mb-1.5">Số lượng tồn</p>
                                        <p className="text-xl font-black text-slate-900 flex items-baseline gap-1.5">
                                          {formatNumber(batch.stock)}
                                          <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">{products.find(p => p.id === batch.productId)?.unit || 'Đơn vị'}</span>
                                        </p>
                                      </div>
                                      {batch.lastExportDate && (
                                         <p className="text-[9px] font-bold text-rose-500 uppercase tracking-tighter mt-1 text-right">Xuất gần nhất: {format(parseISO(batch.lastExportDate), 'dd/MM')}</p>
                                      )}
                                    </div>
                                  </div>
                                ))
                              ) : (
                                <div className="py-20 text-center flex flex-col items-center">
                                  <AlertCircle className="w-12 h-12 text-slate-200 mb-4" />
                                  <p className="text-sm font-black text-slate-400 uppercase tracking-[0.2em] italic">Sản phẩm hiện đã hết hàng tồn kho</p>
                                </div>
                              )}
                            </div>
                          </div>
                          
                          <div className="bg-slate-50 p-6 border-t border-slate-100 flex items-center justify-center gap-3">
                            <div className="flex items-center gap-2">
                              <Info className="w-4 h-4 text-primary" />
                              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest italic">
                                Hệ thống đang áp dụng phương pháp FIFO để trừ tồn kho theo lô
                              </p>
                            </div>
                          </div>
                        </Card>
                      </div>
                    )}
                </div>
              )}

              {activeTab === 'reports' && (
                <div className="space-y-8">
                  {/* Report Sub-Tabs */}
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                    <div className="flex flex-col sm:flex-row items-start sm:items-center gap-4">
                      <div className="flex items-center gap-1.5 p-1.5 bg-white border border-slate-100 rounded-2xl w-fit premium-shadow">
                        {[
                          { id: 'summary', label: 'Tổng hợp', icon: Layers },
                          { id: 'in', label: 'Báo cáo Nhập', icon: PlusCircle },
                          { id: 'out', label: 'Báo cáo Xuất', icon: MinusCircle },
                        ].map(st => (
                          <button
                            key={st.id}
                            onClick={() => setReportSubTab(st.id as any)}
                            className={cn(
                              "flex items-center gap-2 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest transition-all",
                              reportSubTab === st.id 
                                ? "bg-slate-900 text-white shadow-xl shadow-slate-200" 
                                : "text-slate-400 hover:text-slate-900 hover:bg-slate-50"
                            )}
                          >
                            <st.icon className="w-4 h-4" />
                            {st.label}
                          </button>
                        ))}
                      </div>

                      <div className="relative w-full sm:w-64">
                         <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-primary transition-colors">
                           <Search className="w-4 h-4" />
                         </div>
                         <input 
                           type="text"
                           placeholder="Lọc theo Mã lô..."
                           value={batchSearchQuery}
                           onChange={(e) => setBatchSearchQuery(e.target.value)}
                           className="w-full pl-11 pr-4 py-3 bg-white border border-slate-100 rounded-2xl text-xs font-bold focus:ring-4 focus:ring-primary/5 focus:border-primary focus:outline-none premium-shadow transition-all"
                         />
                         {batchSearchQuery && (
                           <button 
                             onClick={() => setBatchSearchQuery('')}
                             className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-rose-500 transition-colors"
                           >
                             <Trash2 className="w-3.5 h-3.5" />
                           </button>
                         )}
                      </div>
                    </div>

                    <div className="flex gap-2">
                       <Button 
                         variant="secondary" 
                         className="px-6 py-2.5 bg-white border border-slate-100"
                         onClick={handleExportReportToExcel}
                       >
                          <Download className="w-4 h-4" /> Xuất Báo Cáo
                       </Button>
                    </div>
                  </div>

                  {batchLifecycle && (
                    <div 
                      className="bg-white border-2 border-slate-900 rounded-[32px] overflow-hidden shadow-2xl shadow-slate-200"
                    >
                      <div className="bg-slate-900 px-8 py-6 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                        <div className="flex items-center gap-4">
                          <div className="w-12 h-12 rounded-2xl bg-amber-400 flex items-center justify-center">
                            <ShieldCheck className="w-6 h-6 text-slate-900" />
                          </div>
                          <div>
                            <h3 className="text-white text-xl font-black uppercase tracking-tight italic font-serif">TRUY XUẤT NGUỒN GỐC LÔ HÀNG</h3>
                            <p className="text-amber-400 text-[10px] font-black uppercase tracking-[0.2em] mt-0.5">Mã lô: {batchLifecycle.batchNumber}</p>
                          </div>
                        </div>
                        <div className="flex gap-3">
                          <div className="px-4 py-2 bg-white/10 rounded-xl backdrop-blur-md border border-white/10">
                            <p className="text-[9px] font-bold text-white/50 uppercase tracking-widest leading-none mb-1">Tồn kho lô</p>
                            <p className="text-sm font-black text-white">{formatNumber(batchLifecycle.balance)} <span className="text-[10px] font-bold text-white/60">{batchLifecycle.unit}</span></p>
                          </div>
                        </div>
                      </div>

                      <div className="p-8 grid grid-cols-1 lg:grid-cols-2 gap-8 bg-slate-50/30">
                        {/* Lịch sử Nhập */}
                        <div className="space-y-4">
                          <div className="flex items-center gap-2 mb-2">
                            <PlusCircle className="w-4 h-4 text-emerald-600" />
                            <h4 className="text-xs font-black text-slate-900 uppercase tracking-widest">Lịch sử Nhập hàng</h4>
                          </div>
                          <div className="space-y-3">
                            {batchLifecycle.imports.length > 0 ? batchLifecycle.imports.map(imp => (
                              <div key={imp.id} className="bg-white p-4 rounded-2xl border border-emerald-100 shadow-sm flex justify-between items-center group hover:border-emerald-300 transition-all">
                                <div className="flex items-center gap-4">
                                  <div className="w-10 h-10 rounded-xl bg-emerald-50 flex items-center justify-center text-emerald-600 font-black text-xs">
                                    {format(parseISO(imp.date), 'dd/MM')}
                                  </div>
                                  <div>
                                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-tighter">Từ: {imp.partnerName}</p>
                                    <h5 className="text-xs font-bold text-slate-900">{imp.productName}</h5>
                                  </div>
                                </div>
                                <div className="text-right">
                                  <p className="text-sm font-black text-emerald-600">+{formatNumber(imp.quantity)}</p>
                                  <p className="text-[9px] font-bold text-slate-400 uppercase">{imp.unit || batchLifecycle.unit}</p>
                                </div>
                              </div>
                            )) : (
                              <div className="py-8 text-center text-slate-400 text-[10px] font-black uppercase tracking-widest italic border-2 border-dashed border-slate-200 rounded-2xl">
                                Không tìm thấy dữ liệu nhập của lô này
                              </div>
                            )}
                          </div>
                        </div>

                        {/* Lịch sử Xuất */}
                        <div className="space-y-4">
                          <div className="flex items-center gap-2 mb-2">
                            <MinusCircle className="w-4 h-4 text-rose-600" />
                            <h4 className="text-xs font-black text-slate-900 uppercase tracking-widest">Lịch sử Xuất hàng</h4>
                          </div>
                          <div className="space-y-3">
                            {batchLifecycle.exports.length > 0 ? batchLifecycle.exports.map(exp => (
                              <div key={exp.id} className="bg-white p-4 rounded-2xl border border-rose-100 shadow-sm flex justify-between items-center group hover:border-rose-300 transition-all">
                                <div className="flex items-center gap-4">
                                  <div className="w-10 h-10 rounded-xl bg-rose-50 flex items-center justify-center text-rose-600 font-black text-xs">
                                    {format(parseISO(exp.date), 'dd/MM')}
                                  </div>
                                  <div>
                                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-tighter">Đến: {exp.partnerName}</p>
                                    <h5 className="text-xs font-bold text-slate-900">{exp.productName}</h5>
                                  </div>
                                </div>
                                <div className="text-right">
                                  <p className="text-sm font-black text-rose-600">-{formatNumber(exp.quantity)}</p>
                                  <p className="text-[9px] font-bold text-slate-400 uppercase">{exp.unit || batchLifecycle.unit}</p>
                                </div>
                              </div>
                            )) : (
                              <div className="py-8 text-center text-slate-400 text-[10px] font-black uppercase tracking-widest italic border-2 border-dashed border-slate-200 rounded-2xl">
                                Lô này chưa phát sinh giao dịch xuất kho
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                      
                      <div className="bg-slate-50 p-4 px-8 border-t border-slate-100 flex items-center gap-4">
                        <AlertCircle className="w-3.5 h-3.5 text-amber-500" />
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest leading-none">
                          Truy xuất dữ liệu trên toàn bộ hệ thống (Bao gồm dữ liệu ngoài khoảng thời gian đang lọc)
                        </p>
                      </div>
                    </div>
                  )}

                  {reportSubTab === 'summary' && (
                    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
                      {/* Strategic "Speaking" Metrics */}
                      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                        <Card className="bg-gradient-to-br from-slate-900 to-slate-800 border-none shadow-xl p-4 sm:p-6 relative overflow-hidden group">
                          <div className="relative z-10">
                            <p className="text-[10px] sm:text-xs font-black uppercase tracking-[0.2em] text-slate-400 mb-1.5">Giá trị tồn kho</p>
                            <h3 className="text-xl sm:text-2xl font-black text-white">{(stats.totalValue / 1000000).toFixed(1)}M <span className="text-xs sm:text-sm font-bold text-slate-500">VNĐ</span></h3>
                            <div className="mt-3 sm:mt-4 flex items-center gap-2 text-[10px] sm:text-xs font-bold text-emerald-400 bg-emerald-400/10 w-fit px-2 sm:px-2.5 py-1 rounded-lg">
                              <TrendingUp className="w-3.5 h-3.5 sm:w-4 sm:h-4" /> +2.4% vs tuần trước
                            </div>
                          </div>
                          <DollarSign className="absolute right-[-5px] bottom-[-5px] sm:right-[-10px] sm:bottom-[-10px] w-16 h-16 sm:w-24 sm:h-24 text-white/5 group-hover:scale-110 transition-transform" />
                        </Card>

                        <Card className="bg-white border border-slate-100 shadow-sm p-4 sm:p-6 relative overflow-hidden group">
                          <p className="text-[10px] sm:text-xs font-black uppercase tracking-[0.2em] text-slate-400 mb-1.5">Điểm sức khỏe kho</p>
                          <div className="flex items-center gap-3">
                            <h3 className="text-xl sm:text-2xl font-black text-slate-900">
                              {Math.round((stats.inventoryHealth.healthy / (inventory.length || 1)) * 100)}%
                            </h3>
                            <div className="flex-1 h-2 sm:h-2.5 bg-slate-100 rounded-full overflow-hidden">
                              <div 
                                className="h-full bg-emerald-500 rounded-full transition-all duration-1000" 
                                style={{ width: `${(stats.inventoryHealth.healthy / (inventory.length || 1)) * 100}%` }} 
                              />
                            </div>
                          </div>
                          <p className="text-[10px] sm:text-xs font-bold text-slate-500 mt-3 sm:mt-4 leading-relaxed">
                            {stats.inventoryHealth.healthy} sản phẩm đạt ngưỡng an toàn.
                          </p>
                        </Card>

                        <Card className="bg-rose-50 border-rose-100 p-4 sm:p-6 relative overflow-hidden group">
                          <p className="text-[10px] sm:text-xs font-black uppercase tracking-[0.2em] text-rose-400 mb-1.5">Cảnh báo cạn kho</p>
                          <div className="flex items-end gap-2">
                             <h3 className="text-2xl sm:text-3xl font-black text-rose-600">{stats.inventoryHealth.low + stats.inventoryHealth.out}</h3>
                             <p className="text-[10px] sm:text-xs font-bold text-rose-400 mb-1 sm:mb-1.5 uppercase">Sản phẩm</p>
                          </div>
                          <div className="mt-2 sm:mt-3 flex gap-2">
                            <div className="px-2 py-0.5 sm:py-1 rounded bg-rose-200 text-rose-700 text-[9px] sm:text-[10px] font-black uppercase">{stats.inventoryHealth.out} Đã hết</div>
                            <div className="px-2 py-0.5 sm:py-1 rounded bg-amber-200 text-amber-700 text-[9px] sm:text-[10px] font-black uppercase">{stats.inventoryHealth.low} Sắp hết</div>
                          </div>
                          <AlertTriangle className="absolute right-[-5px] top-[-5px] w-12 h-12 sm:w-16 sm:h-16 text-rose-500/10" />
                        </Card>

                        <Card className="bg-indigo-50 border-indigo-100 p-4 sm:p-6 relative overflow-hidden group">
                          <p className="text-[10px] sm:text-xs font-black uppercase tracking-[0.2em] text-indigo-400 mb-1.5">Hiệu suất vận hành</p>
                          <h3 className="text-xl sm:text-2xl font-black text-indigo-900">{stats.turnoverRate} <span className="text-xs sm:text-sm">vòng/kỳ</span></h3>
                          <div className="mt-3 sm:mt-4 flex flex-col gap-1 sm:gap-1.5">
                            <div className="flex justify-between text-[9px] sm:text-xs font-bold text-indigo-500 uppercase tracking-tighter">
                              <span>Tốc độ xuất bình quân</span>
                              <span>{formatNumber(stats.totalOut / (timeFilter === 'all' ? 30 : 7))}/ngày</span>
                            </div>
                            <div className="w-full h-1 sm:h-1.5 bg-indigo-200 rounded-full" />
                          </div>
                          <HandCoins className="absolute right-[-5px] bottom-[-5px] sm:right-[-10px] sm:bottom-[-10px] w-16 h-16 sm:w-24 sm:h-24 text-indigo-500/10" />
                        </Card>
                      </div>

                      {/* Revenue Summary Integration */}
                      {filteredRevenueByTime.length > 0 ? (
                        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                           <Card className="lg:col-span-2 bg-slate-900 border-none shadow-2xl p-4 sm:p-8 relative overflow-hidden group">
                                <div className="absolute top-0 right-0 w-64 h-64 bg-primary/5 rounded-full -translate-y-1/2 translate-x-1/2" />
                                <div className="relative z-10 flex flex-col md:flex-row gap-8 items-center">
                                  <div className="space-y-2 text-center md:text-left">
                                    <p className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Tổng doanh thu kỳ này</p>
                                    <h3 className="text-2xl sm:text-4xl font-black text-white">{formatNumber(filteredRevenueByTime.reduce((a, b) => a + b.totalAmount, 0))} <span className="text-[10px] sm:text-sm font-bold opacity-40">VNĐ</span></h3>
                                    <p className="text-xs font-bold text-emerald-400 flex items-center gap-1 justify-center md:justify-start">
                                      <TrendingUp className="w-3 h-3" /> Tăng trưởng ổn định dựa trên báo cáo
                                    </p>
                                  </div>
                                  <div className="hidden md:block w-px h-16 bg-white/10" />
                                  <div className="flex-1 w-full">
                                    <p className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-500 mb-4">Top Đối tác mua nhiều nhất</p>
                                    <div className="space-y-3">
                                      {Object.entries(filteredRevenueByTime.reduce((acc: any, curr) => {
                                        acc[curr.partnerName] = (acc[curr.partnerName] || 0) + curr.totalAmount;
                                        return acc;
                                      }, {})).sort((a: any, b: any) => b[1] - a[1]).slice(0, 2).map(([name, val]) => (
                                        <div key={name} className="flex items-center justify-between">
                                          <span className="text-xs font-bold text-white/70 uppercase tracking-tight">{name}</span>
                                          <span className="text-xs font-black text-white">{formatNumber(val as number)}đ</span>
                                        </div>
                                      ))}
                                    </div>
                                  </div>
                                </div>
                           </Card>
                           <Card className="bg-white border-slate-100 p-4 sm:p-8 flex flex-col justify-center items-center text-center">
                              <p className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400 mb-4">Sản phẩm tiêu thụ chính</p>
                              <div className="w-16 h-16 bg-primary/10 rounded-2xl flex items-center justify-center text-primary mb-3">
                                <Package className="w-8 h-8" />
                              </div>
                              <h4 className="font-black text-slate-900 uppercase">
                                {Object.entries(filteredRevenueByTime.reduce((acc: any, curr) => {
                                  acc[curr.productName] = (acc[curr.productName] || 0) + curr.quantity;
                                  return acc;
                                }, {})).sort((a: any, b: any) => b[1] - a[1])[0][0]}
                              </h4>
                              <p className="text-xs font-bold text-slate-400 mt-1 uppercase tracking-widest">Dẫn đầu sản lượng</p>
                           </Card>
                        </div>
                      ) : (
                        <div className="bg-primary/5 border border-primary/10 rounded-[32px] p-8 flex flex-col md:flex-row items-center justify-between gap-6 group hover:bg-primary/10 transition-all duration-500">
                          <div className="flex items-center gap-6">
                            <div className="w-16 h-16 bg-white rounded-2xl shadow-xl flex items-center justify-center text-primary transform group-hover:rotate-6 transition-transform">
                              <FileSpreadsheet className="w-8 h-8" />
                            </div>
                            <div className="space-y-1 text-center md:text-left">
                              <h4 className="text-lg font-black text-slate-900 uppercase tracking-tight">Kích hoạt Phân tích Doanh thu</h4>
                              <p className="text-sm text-slate-400 font-bold uppercase tracking-wider leading-relaxed max-w-sm">
                                Tải lên tệp Excel của đại lý tại mục <span className="text-primary underline cursor-pointer" onClick={() => setActiveTab('revenue-mgmt')}>Quản lý Doanh thu</span> để tích hợp vào báo cáo tổng hợp.
                              </p>
                            </div>
                          </div>
                          <button 
                            onClick={() => setActiveTab('revenue-mgmt')}
                            className="px-8 py-3.5 bg-primary text-white rounded-2xl font-black text-[10px] uppercase tracking-[0.2em] shadow-xl shadow-primary/20 hover:scale-105 active:scale-95 transition-all"
                          >
                            Đến mục Doanh thu
                          </button>
                        </div>
                      )}

                      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                        <Card title="Phân tích Doanh thu tiềm năng (Hàng tồn)" noPadding>
                          <div className="overflow-x-auto">
                             <table className="w-full text-left">
                               <thead>
                                 <tr className="bg-slate-50/50">
                                   <th className="py-4 px-6 font-bold text-[10px] text-slate-400 uppercase tracking-widest">Sản phẩm</th>
                                   <th className="py-4 px-6 font-bold text-[10px] text-slate-400 uppercase tracking-widest text-right">Tồn hiện tại</th>
                                   <th className="py-4 px-6 font-bold text-[10px] text-slate-400 uppercase tracking-widest text-right">Giá trị (VNĐ)</th>
                                 </tr>
                               </thead>
                               <tbody className="divide-y divide-slate-100">
                                 {inventory.sort((a, b) => b.stock - a.stock).slice(0, 5).map(item => {
                                   const product = products.find(p => p.id === item.productId);
                                   const value = item.stock * (product?.price || 0);
                                   return (
                                     <tr key={item.productId} className="hover:bg-slate-50 transition-colors">
                                       <td className="py-4 px-6 font-bold text-slate-900 text-sm">{item.productName}</td>
                                       <td className="py-4 px-6 text-right font-mono text-sm">{formatNumber(item.stock)}</td>
                                       <td className="py-4 px-6 text-right font-mono text-sm font-black text-primary">{formatNumber(value)}</td>
                                     </tr>
                                   )
                                 })}
                               </tbody>
                             </table>
                          </div>
                        </Card>

                        <Card title="Xu hướng nhập/xuất bia">
                          <div className="h-[300px] w-full mt-4">
                            <ResponsiveContainer width="100%" height="100%">
                              <BarChart data={chartData}>
                                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E2E8F0" />
                                <XAxis dataKey="name" fontSize={12} tickLine={false} axisLine={false} />
                                <YAxis fontSize={12} tickLine={false} axisLine={false} />
                                <Tooltip 
                                  contentStyle={{ 
                                    borderRadius: '12px', 
                                    border: 'none', 
                                    boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.1)',
                                    fontFamily: 'Inter, sans-serif',
                                    fontSize: '12px'
                                  }} 
                                />
                                <Legend wrapperStyle={{ fontSize: '11px', fontWeight: 'Bold', textTransform: 'uppercase', paddingTop: '20px' }} />
                                <Bar dataKey="Nhập" fill="#2563eb" radius={[4, 4, 0, 0]} barSize={20} />
                                <Bar dataKey="Xuất" fill="#f43f5e" radius={[4, 4, 0, 0]} barSize={20} />
                                <Bar dataKey="Doanh thu" fill="#10b981" radius={[4, 4, 0, 0]} barSize={20} />
                              </BarChart>
                            </ResponsiveContainer>
                          </div>
                        </Card>
                      </div>

                      <Card title="Báo cáo Lưu chuyển & Tồn kho" noPadding>
                        <div className="overflow-x-auto">
                          <table className="w-full text-left">
                            <thead>
                              <tr className="bg-slate-50/50">
                                 <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6">Sản phẩm</th>
                                 <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6 text-right">Tồn đầu</th>
                                 <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6 text-right">Nhập</th>
                                 <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6 text-right">Xuất</th>
                                 <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6 text-right">Tồn cuối</th>
                                 <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6 text-right">Tình trạng</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                              {flowSummary.map((item) => {
                                return (
                                  <tr key={item.id} className="hover:bg-slate-50 transition-colors group">
                                    <td className="py-4 px-6">
                                      <p className="font-bold text-slate-900 text-xs">{item.productName}</p>
                                      <p className="text-[10px] text-slate-400 mt-1 uppercase font-bold">{item.category}</p>
                                    </td>
                                    <td className="py-4 px-6 font-mono text-right text-xs font-bold text-slate-500">
                                      {formatNumber(item.openingStock)}
                                    </td>
                                    <td className="py-4 px-6 font-mono text-right text-xs font-bold text-emerald-600">
                                      {formatNumber(item.in)}
                                    </td>
                                    <td className="py-4 px-6 font-mono text-right text-xs font-bold text-rose-500">
                                      {formatNumber(item.out)}
                                    </td>
                                    <td className="py-4 px-6 font-mono text-right text-xs font-black text-slate-900">
                                      {formatNumber(item.closingStock)}
                                    </td>
                                    <td className="py-4 px-6 text-right">
                                       {item.in > 0 && item.out / item.in > 0.8 ? (
                                         <span className="text-[9px] font-black uppercase text-rose-500 bg-rose-50 px-2 py-0.5 rounded border border-rose-100">Dòng vốn nhanh</span>
                                       ) : item.in > 0 && item.out / item.in < 0.2 ? (
                                         <span className="text-[9px] font-black uppercase text-amber-500 bg-amber-50 px-2 py-0.5 rounded border border-amber-100">Tồn đọng</span>
                                       ) : (
                                         <span className="text-[9px] font-black uppercase text-slate-400 bg-slate-50 px-2 py-0.5 rounded border border-slate-100">Ổn định</span>
                                       )}
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      </Card>
                    </div>
                  )}

                  {(reportSubTab === 'in' || reportSubTab === 'out') && (
                    <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
                      <Card title={reportSubTab === 'in' ? "Chi Tiết Nhật Ký Nhập Kho" : "Chi Tiết Nhật Ký Xuất Kho"} noPadding>
                         <div className="overflow-x-auto">
                            <table className="w-full text-left">
                              <thead>
                                <tr className="bg-slate-50/50">
                                  <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6">Ngày thực nhập/xuất</th>
                                  <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6">Sản phẩm</th>
                                  <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6 text-right">Số lượng</th>
                                  <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6">Người thực hiện</th>
                                  <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6">Đối tác</th>
                                  <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6 text-center">Mã lô</th>
                                  <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6">Ghi chú</th>
                                  <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6 text-center">Thao tác</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100">
                                  {filteredTransactionsForReport
                                    .filter(t => reportSubTab === 'in' ? (t.type === 'IN' || t.type === 'OPENING') : t.type === 'OUT')
                                    .map((t, idx, arr) => {
                                      const isGrouped = t.referenceGroupId && (
                                        (idx > 0 && arr[idx-1].referenceGroupId === t.referenceGroupId) ||
                                        (idx < arr.length - 1 && arr[idx+1].referenceGroupId === t.referenceGroupId)
                                      );
                                      return (
                                        <tr 
                                          key={t.id} 
                                          className={cn(
                                            "hover:bg-slate-50 transition-colors relative group/row",
                                            isGrouped ? "bg-amber-50/10" : ""
                                          )}
                                        >
                                          <td className="py-4 px-6 relative">
                                            {isGrouped && (
                                              <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-amber-400" />
                                            )}
                                            <div className="text-[11px] font-bold text-slate-900 font-mono flex items-center gap-2">
                                              {isGrouped && <Layers className="w-3 h-3 text-amber-500" title="Giao dịch xuất nhiều lô" />}
                                              {formatDate(t.date)}
                                            </div>
                                          </td>
                                      <td className="py-4 px-6">
                                        <div className="flex flex-col">
                                          <div className="font-bold text-slate-900 text-sm leading-tight">{t.productName}</div>
                                          <div className="text-[10px] uppercase font-black tracking-widest text-slate-400 mt-0.5">{t.category}</div>
                                          {isGrouped && (
                                            <div className="text-[8px] font-black text-amber-600 uppercase tracking-tighter mt-1 bg-amber-100/50 w-fit px-1 rounded">
                                              Phần của lệnh xuất gộp
                                            </div>
                                          )}
                                        </div>
                                      </td>
                                      <td className="py-4 px-6 text-right">
                                        <span className={cn(
                                          "font-mono font-black text-sm",
                                          t.type === 'IN' ? "text-emerald-600" : "text-rose-600"
                                        )}>
                                          {formatNumber(t.quantity)}
                                        </span>
                                      </td>
                                      <td className="py-4 px-6">
                                        <div className="flex items-center gap-2">
                                          <div className="w-7 h-7 rounded-full bg-slate-100 flex items-center justify-center text-[10px] font-bold text-slate-500 ring-1 ring-slate-200">
                                            {t.createdBy.charAt(0).toUpperCase()}
                                          </div>
                                          <span className="text-xs font-bold text-slate-700">{t.createdBy}</span>
                                        </div>
                                      </td>
                                      <td className="py-4 px-6 text-xs font-bold text-slate-600 uppercase tracking-tight">
                                        {t.partnerName}
                                      </td>
                                      <td className="py-4 px-6 text-center">
                                        <button 
                                          onClick={() => {
                                            if (t.batchNumber) {
                                              setBatchSearchQuery(t.batchNumber);
                                              setReportSubTab(t.type === 'OUT' ? 'out' : 'in');
                                            }
                                          }}
                                          className={cn(
                                            "inline-block px-2 py-1 rounded text-[10px] font-black font-mono transition-all",
                                            t.batchNumber 
                                              ? "bg-slate-100 text-slate-600 hover:bg-slate-900 hover:text-white cursor-pointer" 
                                              : "text-slate-300"
                                          )}
                                        >
                                          {t.batchNumber || '—'}
                                        </button>
                                      </td>
                                      <td className="py-4 px-6">
                                        <div className="flex items-center gap-2 min-w-[100px]">
                                           <span className="text-xs text-slate-400 italic italic">{t.notes || '—'}</span>
                                           {t.evidencePhotoUrl && (
                                              <div className="flex gap-1.5">
                                                <button 
                                                  onClick={() => window.open(t.evidencePhotoUrl, '_blank')}
                                                  className="w-6 h-6 bg-primary/5 text-primary rounded flex items-center justify-center hover:bg-primary hover:text-white transition-all"
                                                  title="Xem ảnh"
                                                >
                                                  <ImageIcon className="w-3 h-3" />
                                                </button>
                                                <a 
                                                  href={t.evidencePhotoUrl} 
                                                  download={`bien-ban-${t.id}.png`}
                                                  className="w-6 h-6 bg-slate-100 text-slate-500 rounded flex items-center justify-center hover:bg-slate-900 hover:text-white transition-all"
                                                  title="Tải về"
                                                >
                                                  <Download className="w-3 h-3" />
                                                </a>
                                              </div>
                                           )}
                                        </div>
                                      </td>
                                      <td className="py-4 px-6 text-center">
                                        <button
                                          onClick={() => handleDeleteTransaction(t.id, t.productName)}
                                          className="p-2 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors"
                                          title="Xóa giao dịch"
                                        >
                                          <Trash2 className="w-4 h-4" />
                                        </button>
                                      </td>
                                    </tr>
                                  ); })
                                }
                                {filteredTransactionsForReport.filter(t => reportSubTab === 'in' ? (t.type === 'IN' || t.type === 'OPENING') : t.type === 'OUT').length === 0 && (
                                  <tr>
                                    <td colSpan={7} className="py-20 text-center text-slate-400 text-sm font-bold uppercase tracking-widest opacity-30">Chưa có dữ liệu giao dịch phù hợp.</td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                         </div>
                      </Card>
                    </div>
                  )}
                </div>
              )}

              {activeTab === 'revenue-mgmt' && (
                    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
                      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                        <div className="space-y-1">
                          <h2 className="text-2xl font-black text-slate-900 tracking-tight">Trung tâm Quản lý Doanh thu</h2>
                          <p className="text-slate-400 text-sm font-bold uppercase tracking-widest leading-none">Phân tích & Theo dõi số liệu kinh doanh</p>
                        </div>
                        <div className="flex gap-2">
                           <input 
                             type="file" 
                             ref={revenueInputRef} 
                             className="hidden" 
                             accept=".xlsx, .xls, .csv" 
                             onChange={handleRevenueReportUpload} 
                           />
                           <button 
                             className="flex items-center gap-2 px-6 py-3 bg-rose-600 text-white rounded-2xl font-black text-xs uppercase tracking-widest shadow-xl shadow-rose-200 hover:scale-105 active:scale-95 transition-all"
                             onClick={clearAllRevenueData}
                           >
                              <Trash2 className="w-4 h-4" /> DỌN DẸP
                           </button>
                           <button 
                             className="flex items-center gap-2 px-6 py-3 bg-white border border-slate-100 text-slate-900 rounded-2xl font-black text-xs uppercase tracking-widest hover:bg-slate-50 transition-all shadow-sm"
                             onClick={handleExportRevenueToExcel}
                           >
                              <Download className="w-4 h-4 text-primary" /> Xuất Excel Doanh thu
                           </button>
                           <button 
                             className="flex items-center gap-2 px-6 py-3 bg-amber-500 text-white rounded-2xl font-black text-xs uppercase tracking-widest shadow-xl shadow-amber-200 hover:scale-105 active:scale-95 transition-all"
                             onClick={downloadRevenueTemplate}
                           >
                              <Download className="w-4 h-4" /> Tải File Mẫu
                           </button>
                           <button 
                             className="flex items-center gap-2 px-6 py-3 bg-primary text-white rounded-2xl font-black text-xs uppercase tracking-widest shadow-xl shadow-primary/20 hover:scale-105 active:scale-95 transition-all"
                             onClick={() => revenueInputRef.current?.click()}
                           >
                              <FileSpreadsheet className="w-4 h-4" /> Nạp Báo cáo Doanh thu (Excel)
                           </button>
                        </div>
                      </div>

                      {revenueData.length > 0 ? (
                        <>
                          {/* Phân tích Doanh thu Chuyên sâu (Dashboard) */}
                          <div className="space-y-6 mb-8">
                            {/* Summary Metrics */}
                            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                              <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm border-l-4 border-l-slate-900">
                                <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Tổng DT Gộp</p>
                                <p className="text-xl font-black text-slate-900">{formatNumber(filteredRevenueByTime.reduce((a, b) => a + b.totalAmount, 0))}</p>
                                <p className="text-[8px] font-bold text-emerald-500 mt-1">Ghi nhận từ HĐ</p>
                              </div>
                              <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm border-l-4 border-l-emerald-500">
                                <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Sản lượng (ĐV)</p>
                                <p className="text-xl font-black text-emerald-600">{formatNumber(filteredRevenueByTime.reduce((a, b) => a + b.quantity, 0))}</p>
                                <p className="text-[8px] font-bold text-slate-300 mt-1">Tổng số đơn vị bán ra</p>
                              </div>
                              <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm border-l-4 border-l-indigo-500">
                                <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Số lượng HĐ</p>
                                <p className="text-xl font-black text-slate-900">{groupedRevenue.length}</p>
                                <p className="text-[8px] font-bold text-slate-300 mt-1">Giao dịch trong kỳ</p>
                              </div>
                              <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm border-l-4 border-l-amber-500">
                                <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Giá trị trung bình</p>
                                <p className="text-xl font-black text-amber-600">{formatNumber(groupedRevenue.length > 0 ? (filteredRevenueByTime.reduce((a, b) => a + b.totalAmount, 0) / groupedRevenue.length) : 0)}</p>
                                <p className="text-[8px] font-bold text-slate-300 mt-1">Bình quân / Hóa đơn</p>
                              </div>
                            </div>
                          </div>



                          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                            <div className="bg-slate-900 rounded-3xl p-6 text-white shadow-xl flex flex-col justify-between">
                              <div className="w-8 h-8 bg-white/10 rounded-lg flex items-center justify-center mb-4">
                                <DollarSign className="w-4 h-4 text-emerald-400" />
                              </div>
                              <div className="space-y-1">
                                <p className="text-[10px] font-black uppercase tracking-[0.2em] text-white/40">Tổng Doanh thu Gộp</p>
                                <h3 className="text-2xl font-black">{formatNumber(filteredRevenueByTime.reduce((a, b) => a + b.totalAmount, 0))} <span className="text-xs">đ</span></h3>
                              </div>
                            </div>
                            
                            <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm flex flex-col justify-between">
                              <div className="w-8 h-8 bg-blue-50 rounded-lg flex items-center justify-center mb-4 font-black text-blue-600 text-[10px]">#</div>
                              <div className="space-y-1">
                                <p className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">Quy mô Danh mục (SKU)</p>
                                <h3 className="text-2xl font-black text-slate-900">{new Set(filteredRevenueByTime.map(r => r.productName)).size} <span className="text-xs text-slate-400">Mã hàng</span></h3>
                              </div>
                            </div>

                            <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm flex flex-col justify-between">
                              <div className="w-8 h-8 bg-indigo-50 rounded-lg flex items-center justify-center mb-4">
                                <Package className="w-4 h-4 text-indigo-500" />
                              </div>
                              <div className="space-y-1">
                                <p className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">Tổng sản lượng tiêu thụ</p>
                                <h3 className="text-2xl font-black text-slate-900">{formatNumber(filteredRevenueByTime.reduce((a, b) => a + b.quantity, 0))} <span className="text-xs text-slate-400">ĐV</span></h3>
                              </div>
                            </div>

                            <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm flex flex-col justify-between">
                              <div className="w-8 h-8 bg-amber-50 rounded-lg flex items-center justify-center mb-4">
                                <ArrowUpRight className="w-4 h-4 text-amber-500" />
                              </div>
                              <div className="space-y-1">
                                <p className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">Đơn hàng Lớn nhất</p>
                                <h3 className="text-2xl font-black text-slate-900">{formatNumber(filteredRevenueByTime.length > 0 ? Math.max(...filteredRevenueByTime.map(r => r.totalAmount)) : 0)} <span className="text-xs text-slate-400">đ</span></h3>
                              </div>
                            </div>
                          </div>

                          <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                            <Card title="Cơ cấu Doanh thu" className="premium-shadow">
                              <div className="h-[250px] sm:h-[380px] w-full mt-4 sm:mt-6">
                                <ResponsiveContainer width="100%" height="100%">
                                  <BarChart 
                                    layout="vertical"
                                    data={
                                      Object.entries(filteredRevenueByTime.reduce((acc: any, curr) => {
                                        acc[curr.productName] = (acc[curr.productName] || 0) + curr.totalAmount;
                                        return acc;
                                      }, {}))
                                      .map(([name, value]) => ({ name, value: value as number }))
                                      .sort((a, b) => b.value - a.value)
                                      .slice(0, 5)
                                    }
                                    margin={{ left: 10, right: 40, top: 10, bottom: 10 }}
                                  >
                                    <CartesianGrid strokeDasharray="3 3" horizontal={true} vertical={false} stroke="#f1f5f9" />
                                    <XAxis type="number" hide />
                                    <YAxis 
                                      dataKey="name" 
                                      type="category" 
                                      fontSize={window.innerWidth < 640 ? 8 : 10} 
                                      width={window.innerWidth < 640 ? 80 : 140} 
                                      axisLine={false} 
                                      tickLine={false}
                                      tick={{ fill: '#64748b', fontWeight: 700 }}
                                    />
                                    <Tooltip 
                                      formatter={(value: any) => formatNumber(value as number) + ' đ'}
                                      cursor={{ fill: '#f8fafc' }}
                                      contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.1)', fontSize: '10px' }}
                                    />
                                    <Bar dataKey="value" fill="#0f172a" radius={[0, 6, 6, 0]} barSize={window.innerWidth < 640 ? 12 : 24} />
                                  </BarChart>
                                </ResponsiveContainer>
                              </div>
                            </Card>

                            <Card title="Khách hàng chiến lược" className="premium-shadow">
                              <div className="h-[250px] sm:h-[380px] w-full mt-4 sm:mt-6">
                                <ResponsiveContainer width="100%" height="100%">
                                  <PieChart>
                                    <Pie
                                      data={
                                        Object.entries(filteredRevenueByTime.reduce((acc: any, curr) => {
                                          acc[curr.partnerName] = (acc[curr.partnerName] || 0) + curr.totalAmount;
                                          return acc;
                                        }, {}))
                                        .map(([name, value]) => ({ name, value: value as number }))
                                        .sort((a, b) => b.value - a.value)
                                        .slice(0, 5)
                                      }
                                      cx="50%"
                                      cy="45%"
                                      innerRadius={window.innerWidth < 640 ? 40 : 70}
                                      outerRadius={window.innerWidth < 640 ? 70 : 110}
                                      paddingAngle={4}
                                      dataKey="value"
                                      stroke="none"
                                    >
                                      {['#0f172a', '#2563eb', '#10b981', '#f59e0b', '#ec4899'].map((color, index) => (
                                        <Cell key={`cell-${index}`} fill={color} />
                                      ))}
                                    </Pie>
                                    <Tooltip 
                                      formatter={(value: any) => formatNumber(value as number) + ' đ'}
                                      contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.1)', fontSize: '10px' }}
                                    />
                                    <Legend iconType="circle" verticalAlign="bottom" height={40} wrapperStyle={{ fontSize: '9px', fontWeight: 800, textTransform: 'uppercase', paddingTop: '10px' }} />
                                  </PieChart>
                                </ResponsiveContainer>
                              </div>
                            </Card>
                          </div>

                          {/* Sổ chi tiết - Chế độ ẩn khi lọc "Tất cả" */}
                          {timeFilter !== 'all' && (
                            <div className="space-y-4 pt-12 border-t border-slate-100">
                              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4">
                                <h4 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
                                  <FileText className="w-5 h-5 text-primary" />
                                  Sổ chi tiết giao dịch từ Báo cáo
                                </h4>
                                
                                <div className="relative group flex-1 max-w-sm">
                                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 group-focus-within:text-primary transition-colors" />
                                  <input 
                                    type="text" 
                                    placeholder="Tìm kiếm đối tác / Đại lý..." 
                                    className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold placeholder:text-slate-300 focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all"
                                    value={revenuePartnerSearch}
                                    onChange={(e) => setRevenuePartnerSearch(e.target.value)}
                                  />
                                  {revenuePartnerSearch && (
                                    <button 
                                      onClick={() => setRevenuePartnerSearch('')}
                                      className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-300 hover:text-slate-500"
                                    >
                                      <X className="w-4 h-4" />
                                    </button>
                                  )}
                                </div>
                              </div>

                              <div className="space-y-3">
                                {groupedRevenue.length > 0 ? groupedRevenue.map((invoice) => {
                                  const isExpanded = expandedInvoices.includes(invoice.invoiceNumber);
                                  return (
                                    <div key={invoice.invoiceNumber} className="group">
                                      <div 
                                        className={cn(
                                          "bg-white border rounded-2xl p-4 sm:p-5 transition-all duration-300 cursor-pointer relative z-10",
                                          isExpanded ? "border-primary shadow-xl shadow-primary/5" : "border-slate-100 hover:border-primary/20 hover:shadow-md"
                                        )}
                                        onClick={() => {
                                          if (invoice.invoiceNumber === 'N/A') return;
                                          setExpandedInvoices(prev => 
                                            prev.includes(invoice.invoiceNumber) ? prev.filter(n => n !== invoice.invoiceNumber) : [...prev, invoice.invoiceNumber]
                                          );
                                        }}
                                      >
                                        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                                          <div className="flex items-center gap-3 sm:gap-4">
                                            <div className={cn(
                                              "w-10 h-10 sm:w-11 sm:h-11 rounded-xl flex items-center justify-center transition-all",
                                              isExpanded ? "bg-primary text-white" : "bg-slate-50 text-slate-400 group-hover:bg-primary/5 group-hover:text-primary"
                                            )}>
                                              <FileText className="w-5 h-5" />
                                            </div>
                                            <div>
                                              <div className="flex items-center gap-2 mb-0.5">
                                                <p className="text-[9px] sm:text-[10px] font-bold text-slate-400 uppercase tracking-widest">{formatDisplayDate(invoice.date)}</p>
                                                {invoice.invoiceNumber !== 'N/A' && (
                                                  <span className="w-0.5 h-0.5 bg-slate-200 rounded-full" />
                                                )}
                                                <span className="text-[8px] sm:text-[9px] font-bold text-primary/60 uppercase tracking-widest">{invoice.items.length} MẶT HÀNG</span>
                                              </div>
                                              <h5 className="text-[12px] sm:text-[13px] font-bold text-slate-800 tracking-tight group-hover:text-primary transition-colors italic">Số HĐ: <span className="not-italic">{invoice.invoiceNumber}</span></h5>
                                            </div>
                                          </div>

                                          <div className="grid grid-cols-2 md:flex items-center md:justify-end gap-x-4 gap-y-2 md:gap-8 w-full md:w-auto px-4 py-3 md:p-0 bg-slate-50/50 md:bg-transparent rounded-xl border border-slate-100 md:border-none">
                                            <div className="text-left md:text-right">
                                              <p className="text-[8px] font-bold text-slate-400 uppercase tracking-widest mb-0.5 leading-none">Đối tác</p>
                                              <p className="text-[10px] sm:text-[11px] font-bold text-slate-600 uppercase truncate max-w-[120px] sm:max-w-none">{invoice.partnerName}</p>
                                            </div>
                                            <div className="text-right">
                                              <p className="text-[8px] font-bold text-slate-400 uppercase tracking-widest mb-0.5 leading-none">Doanh Thu</p>
                                              <p className="text-base sm:text-lg font-bold text-slate-900 leading-none">{formatNumber(invoice.totalAmount)}<span className="text-[9px] font-medium ml-0.5">đ</span></p>
                                            </div>
                                            {invoice.invoiceNumber !== 'N/A' && (
                                              <div className={cn(
                                                "col-span-2 md:col-span-1 flex w-full md:w-7 h-5 md:h-7 rounded-lg md:rounded-full bg-slate-100 md:bg-slate-50 items-center justify-center transition-all mt-1 md:mt-0",
                                                isExpanded ? "bg-primary/10 text-primary md:rotate-180" : "text-slate-300"
                                              )}>
                                                <ChevronDown className={cn("w-4 h-4 transition-transform", isExpanded && "rotate-180 md:rotate-0")} />
                                                <span className="md:hidden text-[8px] font-black uppercase tracking-[0.2em] ml-1">{isExpanded ? 'Thu gọn' : 'Xem chi tiết'}</span>
                                              </div>
                                            )}
                                          </div>
                                        </div>
                                      </div>

                                      {/* Nested Detail View */}
                                      {isExpanded && (
                                        <div 
                                          className="relative -mt-4 pt-8 pb-4 px-8 bg-slate-50/50 border-x border-b border-slate-100/60 rounded-b-2xl z-0 overflow-hidden"
                                        >
                                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                                            {invoice.items.map((item, idx) => (
                                              <div key={idx} className="bg-white border border-slate-100 p-4 rounded-xl flex items-start justify-between shadow-sm">
                                                <div className="space-y-0.5">
                                                  <p className="text-[10px] font-bold text-slate-700 line-clamp-1">{item.productName}</p>
                                                  <p className="text-[9px] font-medium text-slate-400">MÃ: {item.materialCode || 'N/A'}</p>
                                                </div>
                                                <div className="text-right">
                                                  <p className="text-[10px] font-bold text-slate-900">{formatNumber(item.quantity)} {item.unit || products.find(p => p.name === item.productName)?.unit || 'ĐV'}</p>
                                                  <p className="text-[9px] font-bold text-emerald-600">{formatNumber(item.totalAmount)} đ</p>
                                                </div>
                                              </div>
                                            ))}
                                          </div>
                                        </div>
                                      )}
                                    </div>
                                  );
                                }) : (
                                  <div className="py-20 text-center">
                                    <Search className="w-10 h-10 text-slate-100 mx-auto mb-4" />
                                    <p className="text-xs font-bold text-slate-300 uppercase tracking-widest">Không có dữ liệu hóa đơn</p>
                                  </div>
                                )}
                              </div>
                            </div>
                          )}
                        </>
                      ) : (
                        <Card className="py-32 flex flex-col items-center justify-center text-center space-y-4 border-dashed border-2 bg-slate-50/50 rounded-[32px]">
                           <div className="w-20 h-20 bg-white rounded-3xl flex items-center justify-center shadow-xl text-slate-300">
                              <FileSpreadsheet className="w-10 h-10" />
                           </div>
                           <div className="space-y-2">
                             <h4 className="text-xl font-black text-slate-900">Sẵn sàng phân tích doanh thu</h4>
                             <p className="text-sm text-slate-400 max-w-sm mx-auto font-bold uppercase tracking-wider">Chọn file Excel báo cáo từ đại lý để bắt đầu quy trình đối soát.</p>
                           </div>
                           <div className="flex gap-4 mt-4">
                             <button 
                               className="px-8 py-3 bg-primary text-white rounded-2xl font-black text-xs uppercase tracking-widest shadow-xl shadow-primary/20 hover:scale-105 active:scale-95 transition-all"
                               onClick={() => revenueInputRef.current?.click()}
                             >
                               Tải lên hồ sơ doanh thu
                             </button>
                             <button 
                               className="px-8 py-3 bg-amber-500 text-white rounded-2xl font-black text-xs uppercase tracking-widest shadow-xl shadow-amber-200 hover:scale-105 active:scale-95 transition-all"
                               onClick={downloadRevenueTemplate}
                             >
                               Tải File Excel Mẫu
                             </button>
                           </div>
                        </Card>
                      )}
                    </div>
                  )}

              {activeTab === 'in-transit' && (
                <div className="space-y-6">
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                    <div>
                      <h2 className="text-2xl font-black text-slate-900 uppercase tracking-tight">Quản lý đơn đi đường</h2>
                      <p className="text-xs text-slate-400 font-bold uppercase tracking-widest mt-1">Theo dõi và xác nhận các đơn hàng đang vận chuyển (2-3 ngày).</p>
                    </div>
                  </div>

                  <Card noPadding className="premium-shadow overflow-hidden">
                    <div className="overflow-x-auto">
                      <table className="w-full text-left">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-100">
                            <th className="py-4 px-6 font-black text-[10px] text-slate-400 uppercase tracking-widest">Ngày xuất</th>
                            <th className="py-4 px-6 font-black text-[10px] text-slate-400 uppercase tracking-widest">Sản phẩm</th>
                            <th className="py-4 px-6 font-black text-[10px] text-slate-400 uppercase tracking-widest">Đối tác</th>
                            <th className="py-4 px-6 font-black text-[10px] text-slate-400 uppercase tracking-widest text-right">SL Gửi</th>
                            <th className="py-4 px-6 font-black text-[10px] text-slate-400 uppercase tracking-widest text-center">Thao tác</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {transactions.filter(t => t.status === 'in_transit').length > 0 ? (
                            transactions.filter(t => t.status === 'in_transit').map((t) => (
                              <tr key={t.id} className="hover:bg-amber-50/30 transition-all group">
                                <td className="py-4 px-6 text-[11px] font-mono font-bold text-slate-500">
                                  {formatDate(t.date)}
                                </td>
                                <td className="py-4 px-6">
                                  <div className="font-bold text-slate-900 text-sm leading-tight">{t.productName}</div>
                                  <div className="text-[10px] text-slate-400 font-bold uppercase tracking-tighter mt-1">LOT: {t.batchNumber || 'N/A'}</div>
                                </td>
                                <td className="py-4 px-6">
                                  <span className="px-2 py-0.5 bg-white border border-slate-200 rounded-full text-[10px] font-black uppercase tracking-tighter text-slate-700">
                                    {t.partnerName}
                                  </span>
                                </td>
                                <td className="py-4 px-6 text-right font-mono font-black text-sm text-amber-600">
                                  {formatNumber(t.quantity)}
                                </td>
                                <td className="py-4 px-6">
                                  <div className="flex items-center justify-center gap-2">
                                    <button 
                                      onClick={() => {
                                        setSelectedInTransit(t);
                                        setActualReceivedQty(t.quantity);
                                        setShowLossModal(true);
                                      }}
                                      className="px-3 py-1.5 bg-emerald-500 text-white rounded-lg text-[10px] font-black uppercase tracking-widest hover:bg-emerald-600 transition-all flex items-center gap-1.5 shadow-lg shadow-emerald-500/20"
                                    >
                                      <CheckCircle className="w-3.5 h-3.5" />
                                      Xác nhận nhận hàng
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={5} className="py-20 text-center text-slate-300 font-bold uppercase tracking-[0.2em] text-xs">Không có đơn hàng nào đang đi đường.</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </Card>

                  {/* Loss Reporting Modal */}
                    {showLossModal && selectedInTransit && (
                      <div className="fixed inset-0 z-[70] flex items-center justify-center p-4">
                        <div 
                          onClick={() => setShowLossModal(false)}
                          className="absolute inset-0 bg-slate-900/60 backdrop-blur-md"
                        />
                        <div 
                          className="relative w-full max-w-lg bg-white rounded-[32px] shadow-2xl overflow-hidden p-8"
                        >
                          <div className="flex items-center justify-between mb-8">
                            <div>
                              <h3 className="text-xl font-black text-slate-900 uppercase">XÁC NHẬN HOÀN THÀNH ĐƠN HÀNG</h3>
                              <p className="text-[10px] font-black text-slate-400 mt-1 uppercase tracking-widest">Đơn hàng: {selectedInTransit.id}</p>
                            </div>
                            <button onClick={() => {
                              setShowLossModal(false);
                              setLossEvidencePhoto('');
                            }} className="p-2 hover:bg-slate-100 rounded-full transition-colors">
                              <X className="w-5 h-5 text-slate-400" />
                            </button>
                          </div>

                          <div className="space-y-6">
                            <div className="p-4 bg-slate-50 rounded-2xl border border-slate-100 space-y-2">
                                <div className="flex justify-between text-xs">
                                  <span className="text-slate-400 font-bold">Sản phẩm:</span>
                                  <span className="text-slate-900 font-black">{selectedInTransit.productName}</span>
                                </div>
                                <div className="flex justify-between text-xs">
                                  <span className="text-slate-400 font-bold">Số lượng gửi:</span>
                                  <span className="text-amber-600 font-black">{formatNumber(selectedInTransit.quantity)}</span>
                                </div>
                            </div>

                            <Input 
                              label="Số lượng thực nhận (Hàng lành lặn)"
                              type="number"
                              value={actualReceivedQty}
                              onChange={(e: any) => setActualReceivedQty(Number(e.target.value))}
                            />

                            {actualReceivedQty < selectedInTransit.quantity && (
                              <div className="space-y-2">
                                <label className="text-[11px] font-extrabold text-slate-500 uppercase tracking-widest ml-1">Lý do hao hụt / Hư hại</label>
                                <textarea 
                                  className="w-full p-4 bg-slate-50 border border-slate-100 rounded-2xl focus:ring-4 focus:ring-primary/5 focus:border-primary outline-none text-sm font-medium min-h-[100px]"
                                  placeholder="Ví dụ: Nổ lon do va đập, thiếu hàng khi kiểm đếm..."
                                  value={lossReason}
                                  onChange={(e) => setLossReason(e.target.value)}
                                />
                              </div>
                            )}

                            <div className="space-y-3">
                              <label className="text-[11px] font-extrabold text-slate-500 uppercase tracking-widest ml-1">Minh chứng thực tế (Ảnh chụp nhận hàng)</label>
                              
                              {!lossEvidencePhoto ? (
                                <div className="grid grid-cols-2 gap-4">
                                  <label className="flex flex-col items-center justify-center gap-3 p-8 border-2 border-dashed border-slate-200 rounded-3xl hover:border-primary hover:bg-primary/5 transition-all cursor-pointer group">
                                    <Camera className="w-8 h-8 text-slate-300 group-hover:text-primary transition-colors" />
                                    <span className="text-[11px] font-black text-slate-400 uppercase group-hover:text-primary tracking-tighter">Chụp ảnh</span>
                                    <input 
                                      type="file" 
                                      className="hidden" 
                                      accept="image/*" 
                                      capture="environment"
                                      onChange={(e) => {
                                        const file = e.target.files?.[0];
                                        if (file) {
                                          const reader = new FileReader();
                                          reader.onload = async () => {
                                            const compressed = await compressImage(reader.result as string);
                                            setLossEvidencePhoto(compressed);
                                          };
                                          reader.readAsDataURL(file);
                                        }
                                      }}
                                    />
                                  </label>
                                  <label className="flex flex-col items-center justify-center gap-3 p-8 border-2 border-dashed border-slate-200 rounded-3xl hover:border-emerald-500 hover:bg-emerald-50 transition-all cursor-pointer group">
                                    <ImageIcon className="w-8 h-8 text-slate-300 group-hover:text-emerald-500 transition-colors" />
                                    <span className="text-[11px] font-black text-slate-400 uppercase group-hover:text-emerald-500 tracking-tighter">Chọn từ máy</span>
                                    <input 
                                      type="file" 
                                      className="hidden" 
                                      accept="image/*" 
                                      onChange={(e) => {
                                        const file = e.target.files?.[0];
                                        if (file) {
                                          const reader = new FileReader();
                                          reader.onload = async () => {
                                            const compressed = await compressImage(reader.result as string);
                                            setLossEvidencePhoto(compressed);
                                          };
                                          reader.readAsDataURL(file);
                                        }
                                      }}
                                    />
                                  </label>
                                </div>
                              ) : (
                                <div className="relative rounded-2xl overflow-hidden group">
                                  <img src={lossEvidencePhoto} alt="Evidence" className="w-full aspect-video object-cover" />
                                  <div className="absolute inset-0 bg-slate-900/40 opacity-0 group-hover:opacity-100 transition-all flex items-center justify-center">
                                    <button 
                                      onClick={() => setLossEvidencePhoto('')}
                                      className="p-3 bg-rose-500 text-white rounded-full shadow-lg"
                                    >
                                      <Trash2 className="w-5 h-5" />
                                    </button>
                                  </div>
                                </div>
                              )}
                            </div>

                            {actualReceivedQty < selectedInTransit.quantity && (
                              <div className="bg-rose-50 border border-rose-100 p-4 rounded-xl">
                                <div className="flex items-center gap-2 text-rose-600 mb-1">
                                  <AlertCircle className="w-4 h-4" />
                                  <span className="text-[10px] font-black uppercase tracking-widest">Cảnh báo chênh lệch</span>
                                </div>
                                <p className="text-xs text-rose-500 font-bold italic">
                                  Hệ thống sẽ tự động tách {formatNumber(selectedInTransit.quantity - actualReceivedQty)} đơn vị vào danh mục HAO HỤT/HƯ HỎNG.
                                </p>
                              </div>
                            )}

                            <div className="pt-4 flex gap-3">
                              <Button variant="outline" className="flex-1" onClick={() => {
                                setShowLossModal(false);
                                setLossEvidencePhoto('');
                              }}>Hủy</Button>
                              <Button 
                                className={`flex-[2] ${actualReceivedQty < selectedInTransit.quantity ? 'bg-amber-500 hover:bg-amber-600' : 'bg-emerald-500 hover:bg-emerald-600'}`} 
                                onClick={handleReportLoss}
                              >
                                {actualReceivedQty < selectedInTransit.quantity ? 'Xác nhận điều chỉnh' : 'Hoàn thành ghi nhận'}
                              </Button>
                            </div>
                          </div>
                        </div>
                      </div>
                    )}
                </div>
              )}

              {(activeTab === 'import' || activeTab === 'export') && (
                <div className="max-w-4xl mx-auto space-y-6">
                  <div className="text-center space-y-2 mb-8">
                    <div className={cn(
                      "w-16 h-16 mx-auto rounded-2xl flex items-center justify-center mb-4 shadow-xl",
                      activeTab === 'import' ? "bg-blue-600 shadow-blue-200" : "bg-rose-600 shadow-rose-200"
                    )}>
                      {activeTab === 'import' ? <PlusCircle className="text-white w-8 h-8" /> : <MinusCircle className="text-white w-8 h-8" />}
                    </div>
                    <h2 className="text-2xl font-black text-slate-900 italic font-serif">
                      GHI NHẬN {activeTab === 'import' 
                        ? (newTransaction.type === 'OPENING' ? 'TỒN ĐẦU KỲ' : 'NHẬP KHO') 
                        : 'XUẤT KHO'}
                    </h2>
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-[0.3em]">Hệ thống ghi nhận đa mặt hàng (Batch Processing)</p>
                  </div>

                  <Card className="p-4 sm:p-10 relative overflow-hidden">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-8 relative z-10">
                      {activeTab === 'import' && (
                        <div className="md:col-span-2">
                          <Select 
                            label="Hình thức ghi nhận"
                            options={[
                              { value: 'IN', label: 'Nhập kho (Nhập hàng mới về)' },
                              { value: 'OPENING', label: 'Tồn đầu kỳ (Số liệu gốc ban đầu)' }
                            ]}
                            value={newTransaction.type}
                            onChange={(e: any) => {
                              const type = e.target.value as TransactionType;
                              setNewTransaction({ 
                                ...newTransaction, 
                                type,
                                items: newTransaction.items.map(item => ({
                                  ...item,
                                  batchNumber: type === 'OPENING' ? 'Tồn 25/4' : (item.batchNumber === 'Tồn 25/4' ? '' : item.batchNumber)
                                }))
                              });
                            }}
                          />
                        </div>
                      )}
                      
                      <div className="md:col-span-1">
                        <Input 
                          label={activeTab === 'import' 
                            ? (newTransaction.type === 'OPENING' ? 'Ngày chốt tồn đầu' : 'Ngày thực nhập') 
                            : 'Ngày thực xuất'
                          }
                          type="date"
                          value={newTransaction.date}
                          onChange={(e: any) => setNewTransaction({ ...newTransaction, date: e.target.value })}
                        />
                      </div>

                      <div className="md:col-span-1">
                        {newTransaction.type !== 'OPENING' && (
                          <Select 
                            label={activeTab === 'import' ? "Nhà cung cấp" : "Đơn vị nhận (Đối tác)"}
                            options={partners
                              .filter(p => activeTab === 'import' ? p.type === 'SUPPLIER' : p.type !== 'SUPPLIER')
                              .map(p => ({ value: p.id, label: `${p.name} ${p.sapCode ? '[' + p.sapCode + ']' : ''}` }))
                            }
                            value={newTransaction.partnerId}
                            onChange={(e: any) => setNewTransaction({ ...newTransaction, partnerId: e.target.value })}
                          />
                        )}
                      </div>

                      {/* Items List */}
                      <div className="md:col-span-2 mt-4">
                        <div className="flex items-center justify-between mb-4 border-b border-slate-100 pb-2">
                          <h4 className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Danh sách mặt hàng ({newTransaction.items.length})</h4>
                          <div className="flex gap-2">
                            <button 
                              onClick={() => {
                                const missingProducts = products.filter(p => !newTransaction.items.some(item => item.productId === p.id));
                                if (!missingProducts.length) return showNotification("Anh đã chọn hết các mặt hàng hiện có rồi ạ!");
                                
                                setNewTransaction(prev => ({
                                  ...prev,
                                  items: [
                                    ...prev.items.filter(i => i.productId !== ''),
                                    ...missingProducts.map(p => ({ productId: p.id, quantity: 0, batchNumber: '' }))
                                  ]
                                }));
                              }}
                              className="text-[10px] font-black text-amber-600 bg-amber-50 uppercase tracking-widest flex items-center gap-1.5 hover:bg-amber-100 px-3 py-1.5 rounded-lg transition-all"
                            >
                              <Layers className="w-3.5 h-3.5" /> Thêm toàn bộ danh mục
                            </button>
                            <button 
                              onClick={() => {
                                // Simple quick select: allow user to select multiple products at once
                                const availableProducts = products.filter(p => 
                                  !newTransaction.items.some(item => item.productId === p.id)
                                );
                                if (availableProducts.length === 0) {
                                  showNotification("Anh đã chọn hết các mặt hàng hiện có rồi ạ!");
                                  return;
                                }
                                
                                const selectedIds = window.prompt(
                                  "Nhập tên sản phẩm để chọn nhanh (Ngăn cách bởi dấu phẩy):\n" + 
                                  availableProducts.map(p => p.name).join(', ')
                                );
                                
                                if (selectedIds) {
                                  const searchTerms = selectedIds.split(',').map(s => s.trim().toLowerCase());
                                  const matches = availableProducts.filter(p => 
                                    searchTerms.some(term => p.name.toLowerCase().includes(term))
                                  );
                                  
                                  if (matches.length > 0) {
                                    setNewTransaction(prev => ({
                                      ...prev,
                                      items: [
                                        ...prev.items.filter(i => i.productId !== ''),
                                        ...matches.map(p => ({ productId: p.id, quantity: 0, batchNumber: '' }))
                                      ]
                                    }));
                                  }
                                }
                              }}
                              className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-1.5 hover:bg-slate-100 px-3 py-1.5 rounded-lg transition-all"
                            >
                              <Search className="w-3.5 h-3.5" /> Chọn nhanh
                            </button>
                            <button 
                              onClick={addTransactionItem}
                              className="text-[10px] font-black text-primary uppercase tracking-widest flex items-center gap-1.5 hover:bg-primary/5 px-3 py-1.5 rounded-lg transition-all"
                            >
                              <PlusCircle className="w-3.5 h-3.5" /> Thêm dòng
                            </button>
                          </div>
                        </div>

                        <div className="space-y-4">
                            {newTransaction.items.map((item, index) => (
                              <div 
                                key={`item-${index}-${item.productId}`}
                                className="p-4 sm:p-6 bg-slate-50/50 rounded-2xl border border-dotted border-slate-200 relative group/item"
                              >
                                {newTransaction.items.length > 1 && (
                                  <button 
                                    onClick={() => removeTransactionItem(index)}
                                    className="absolute -top-2 -right-2 w-7 h-7 bg-white border border-slate-100 text-rose-500 rounded-full flex items-center justify-center shadow-sm opacity-0 group-hover/item:opacity-100 transition-all hover:bg-rose-50"
                                  >
                                    <X className="w-4 h-4" />
                                  </button>
                                )}
                                
                                <div className="grid grid-cols-1 sm:grid-cols-12 gap-4 items-end">
                                  <div className="sm:col-span-5">
                                    <Select 
                                      label={`Sản phẩm ${index + 1}`}
                                      options={products.map(p => ({ value: p.id, label: `${p.name} (${p.category})` }))}
                                      value={item.productId}
                                      onChange={(e: any) => updateTransactionItem(index, { productId: e.target.value })}
                                    />
                                  </div>
                                  <div className="sm:col-span-3">
                                    <Input 
                                      label={`Số lượng (${products.find(p => p.id === item.productId)?.unit || '—'})`}
                                      type="number"
                                      placeholder="0"
                                      value={item.quantity === 0 ? '' : item.quantity}
                                      onChange={(e: any) => updateTransactionItem(index, { quantity: Number(e.target.value) })}
                                    />
                                  </div>
                                  <div className="sm:col-span-4">
                                    {(activeTab === 'import' || newTransaction.type === 'OPENING') && (
                                      <Input 
                                        label="Số lô (Mã lô nhập)"
                                        placeholder="LOT-XXX"
                                        value={item.batchNumber}
                                        onChange={(e: any) => updateTransactionItem(index, { batchNumber: e.target.value })}
                                      />
                                    )}
                                    {activeTab === 'export' && (
                                      <div className="text-[10px] font-mono font-black text-rose-500 bg-rose-50/50 px-3 py-3.5 rounded-xl border border-rose-100">
                                        FIFO: {item.batchNumber || 'TỰ ĐỘNG'}
                                      </div>
                                    )}
                                  </div>
                                </div>
                              </div>
                            ))}
                        </div>
                      </div>

                      <div className="md:col-span-2">
                        <Input 
                          label="Ghi chú tổng quát (Tùy chọn)"
                          placeholder="Ví dụ: Theo xe tải số hiệu..., Ghi chú chung cho cả đơn..."
                          value={newTransaction.notes}
                          onChange={(e: any) => setNewTransaction({ ...newTransaction, notes: e.target.value })}
                        />
                      </div>

                      {activeTab === 'export' && (
                        <div className="md:col-span-2">
                          <label className="flex items-center gap-3 cursor-pointer p-5 bg-amber-50/50 border border-amber-200/50 rounded-2xl hover:bg-amber-100/50 transition-all">
                            <input 
                              type="checkbox" 
                              className="w-5 h-5 rounded border-amber-300 text-amber-600 focus:ring-amber-500"
                              checked={newTransaction.isInTransit}
                              onChange={(e) => setNewTransaction({ ...newTransaction, isInTransit: e.target.checked })}
                            />
                            <div className="flex flex-col">
                              <span className="text-xs font-black text-amber-900 uppercase tracking-widest">Đơn hàng đi đường (Giao sau)</span>
                              <span className="text-[10px] text-amber-700/70 font-bold">Hàng đang vận chuyển, chưa đến tay đối tác ngay.</span>
                            </div>
                            <Truck className="w-6 h-6 text-amber-500 ml-auto" />
                          </label>
                        </div>
                      )}

                      <div className="md:col-span-2 space-y-4">
                        <label className="text-[11px] font-extrabold text-slate-500 uppercase tracking-widest ml-1">Minh chứng (Ảnh chụp chứng từ cho cả đơn)</label>
                        <div className="flex flex-col gap-4">
                          {!newTransaction.evidencePhotoUrl ? (
                            <div className="flex gap-4">
                              <button 
                                type="button"
                                onClick={() => document.getElementById('photo-upload-multi')?.click()}
                                className="flex-1 h-32 border-2 border-dashed border-slate-200 rounded-3xl flex flex-col items-center justify-center gap-3 text-slate-400 hover:border-primary hover:text-primary hover:bg-primary/5 transition-all group"
                              >
                                <div className="w-12 h-12 bg-slate-50 rounded-2xl flex items-center justify-center group-hover:bg-white transition-colors">
                                  <ImageIcon className="w-6 h-6 group-hover:scale-110 transition-transform" />
                                </div>
                                <span className="text-[10px] font-black uppercase tracking-widest">Tải ảnh lên</span>
                              </button>
                              <button 
                                type="button"
                                onClick={() => document.getElementById('camera-capture-multi')?.click()}
                                className="flex-1 h-32 border-2 border-dashed border-slate-200 rounded-3xl flex flex-col items-center justify-center gap-3 text-slate-400 hover:border-emerald-400 hover:text-emerald-500 hover:bg-emerald-50 transition-all group"
                              >
                                <div className="w-12 h-12 bg-slate-50 rounded-2xl flex items-center justify-center group-hover:bg-white transition-colors">
                                  <Camera className="w-6 h-6 group-hover:scale-110 transition-transform" />
                                </div>
                                <span className="text-[10px] font-black uppercase tracking-widest">Chụp bằng máy ảnh</span>
                              </button>
                            </div>
                          ) : (
                            <div className="relative group rounded-3xl overflow-hidden border-4 border-white shadow-2xl">
                              <img src={newTransaction.evidencePhotoUrl} alt="Preview" className="w-full h-64 object-cover" />
                              <div className="absolute inset-x-0 bottom-0 bg-gradient-to-t from-black/80 to-transparent p-6 flex items-center justify-between">
                                <span className="text-white text-[10px] font-black uppercase tracking-widest">Biên bản / Chứng từ gốc</span>
                                <button 
                                  onClick={() => setNewTransaction({ ...newTransaction, evidencePhotoUrl: '' })}
                                  className="w-10 h-10 bg-rose-500 text-white rounded-xl flex items-center justify-center hover:bg-rose-600 transition-all shadow-lg"
                                >
                                  <X className="w-5 h-5" />
                                </button>
                              </div>
                            </div>
                          )}
                          
                          <input id="photo-upload-multi" type="file" accept="image/*" className="hidden" 
                            onChange={(e) => {
                              const file = e.target.files?.[0];
                              if (file) {
                                const reader = new FileReader();
                                reader.onloadend = async () => {
                                  try {
                                    const compressed = await compressImage(reader.result as string);
                                    setNewTransaction(prev => ({ ...prev, evidencePhotoUrl: compressed }));
                                  } catch (err) {
                                    console.error("Compression failed", err);
                                    setNewTransaction(prev => ({ ...prev, evidencePhotoUrl: reader.result as string }));
                                  }
                                };
                                reader.readAsDataURL(file);
                              }
                            }}
                          />
                          <input id="camera-capture-multi" type="file" accept="image/*" capture="environment" className="hidden" 
                            onChange={(e) => {
                              const file = e.target.files?.[0];
                              if (file) {
                                const reader = new FileReader();
                                reader.onloadend = async () => {
                                  try {
                                    const compressed = await compressImage(reader.result as string);
                                    setNewTransaction(prev => ({ ...prev, evidencePhotoUrl: compressed }));
                                  } catch (err) {
                                    console.error("Compression failed", err);
                                    setNewTransaction(prev => ({ ...prev, evidencePhotoUrl: reader.result as string }));
                                  }
                                };
                                reader.readAsDataURL(file);
                              }
                            }}
                          />
                        </div>
                      </div>
                    </div>
                    
                    <div className="mt-12 flex flex-col sm:flex-row gap-4 items-center border-t border-slate-100 pt-10">
                      <Button 
                        variant={activeTab === 'import' ? 'primary' : 'danger'} 
                        className="w-full sm:flex-1 py-5 text-base shadow-2xl relative group overflow-hidden"
                        onClick={() => handleAddTransaction(activeTab === 'import' ? newTransaction.type : 'OUT')}
                        disabled={loading || newTransaction.items.some(i => !i.productId || i.quantity <= 0)}
                      >
                        <div className="absolute inset-0 bg-white/10 opacity-0 group-hover:opacity-100 transition-opacity" />
                        {loading ? <RefreshCw className="w-5 h-5 animate-spin" /> : (activeTab === 'import' ? <PlusCircle className="w-5 h-5" /> : <MinusCircle className="w-5 h-5" />)}
                        <span className="relative z-10">
                          {activeTab === 'import' 
                            ? (newTransaction.type === 'OPENING' ? 'GHI NHẬN TỒN ĐẦU' : `HOÀN TẤT NHẬP ${newTransaction.items.length} MÓN`) 
                            : `HOÀN TẤT XUẤT ${newTransaction.items.length} MÓN`}
                        </span>
                      </Button>
                      <Button variant="secondary" className="w-full sm:w-auto px-10 py-5" onClick={() => setActiveTab('dashboard')}>Hủy bỏ</Button>
                    </div>

                    <div className="absolute top-0 right-0 w-64 h-64 bg-slate-50 rounded-full -mr-32 -mt-32 opacity-20 pointer-events-none" />
                    <div className="absolute bottom-0 left-0 w-48 h-48 bg-slate-50 rounded-full -ml-24 -mb-24 opacity-20 pointer-events-none" />
                  </Card>
                </div>
              )}

              {activeTab === 'history' && (
                <div className="space-y-6">
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                    <div className="relative max-w-md w-full group">
                      <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 group-focus-within:text-primary transition-colors hover:text-primary" />
                      <input 
                        placeholder="Tìm kiếm lịch sử (Sản phẩm, Đối tác, Số lô...)" 
                        className="w-full pl-11 pr-4 py-3 bg-white border border-slate-100 rounded-2xl focus:ring-4 focus:ring-primary/5 focus:border-primary outline-none text-sm transition-all premium-shadow"
                        value={historySearchQuery}
                        onChange={(e) => setHistorySearchQuery(e.target.value)}
                      />
                    </div>
                    <div className="flex gap-2">
                       <Button variant="secondary" className="bg-white border border-slate-100">
                          <Download className="w-4 h-4" /> Xuất dữ liệu
                       </Button>
                    </div>
                  </div>

                  <Card noPadding>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left">
                        <thead>
                          <tr className="bg-slate-50/50">
                             <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6">Ngày thực nhập/xuất</th>
                             <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6 text-center">Hoạt động</th>
                             <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6">Sản phẩm</th>
                             <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6 text-right">Số lượng</th>
                             <th className="font-bold text-[10px] text-slate-400 uppercase tracking-widest py-4 px-6">Ghi chú & Minh chứng</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {filteredTransactions.length > 0 ? filteredTransactions.map((t) => (
                            <tr key={t.id} className="hover:bg-slate-50 transition-colors group">
                              <td className="py-4 px-6">
                                <div className="text-[11px] font-bold text-slate-500 font-mono">
                                  {formatDate(t.date).split(' ')[0]}
                                </div>
                                <div className="text-[10px] text-slate-400 font-bold">
                                  {formatDate(t.date).split(' ')[1]}
                                </div>
                              </td>
                              <td className="py-4 px-6 text-center">
                                <span className={cn(
                                  "inline-flex items-center justify-center w-8 h-8 rounded-lg",
                                  (t.type === 'IN' || t.type === 'OPENING') ? "bg-emerald-50 text-emerald-600" : 
                                  (t.type === 'LOSS' || t.type === 'DAMAGE') ? "bg-rose-100 text-rose-700" :
                                  "bg-rose-50 text-rose-600"
                                )}>
                                  {(t.type === 'IN' || t.type === 'OPENING') ? <ArrowDownLeft className="w-4 h-4" /> : <ArrowUpRight className="w-4 h-4" />}
                                </span>
                              </td>
                              <td className="py-4 px-6">
                                <div className="font-bold text-slate-900 text-sm leading-tight">
                                  {t.productName}
                                  {t.status === 'in_transit' && (
                                    <span className="ml-2 px-2 py-0.5 bg-amber-100 text-amber-600 rounded-full text-[9px] font-black uppercase tracking-tighter flex inline-flex items-center gap-1 mt-1 md:mt-0">
                                      <Truck className="w-3 h-3" />
                                      Đang đi đường
                                    </span>
                                  )}
                                  {(t.type === 'LOSS' || t.type === 'DAMAGE') && (
                                    <span className="ml-2 px-2 py-0.5 bg-rose-100 text-rose-600 rounded-full text-[9px] font-black uppercase tracking-tighter mt-1 md:mt-0">
                                      Hao hụt / Hư hại
                                    </span>
                                  )}
                                </div>
                                <div className="text-[10px] uppercase font-black tracking-widest text-slate-400 mt-0.5">{t.category} • {t.partnerName}</div>
                                {t.batchNumber && (
                                  <div className="mt-1 text-[10px] font-mono font-black text-primary/60">LOT: {t.batchNumber}</div>
                                )}
                              </td>
                              <td className="py-4 px-6 text-right">
                                <span className={cn(
                                  "font-mono font-black text-sm",
                                  (t.type === 'IN' || t.type === 'OPENING') ? "text-emerald-600" : "text-rose-600"
                                )}>
                                  {(t.type === 'IN' || t.type === 'OPENING') ? '+' : '-'}{formatNumber(t.quantity)}
                                </span>
                              </td>
                              <td className="py-4 px-6">
                                <div className="flex items-center gap-3">
                                  <span className="text-xs text-slate-400 italic max-w-[120px] truncate">{t.notes || '—'}</span>
                                  {t.evidencePhotoUrl && (
                                    <div className="flex gap-1.5">
                                      <button 
                                        onClick={() => window.open(t.evidencePhotoUrl, '_blank')}
                                        className="w-7 h-7 bg-primary/5 text-primary rounded-lg flex items-center justify-center hover:bg-primary hover:text-white transition-all shadow-sm"
                                        title="Xem biên bản"
                                      >
                                        <ImageIcon className="w-4 h-4" />
                                      </button>
                                      <a 
                                        href={t.evidencePhotoUrl} 
                                        download={`bien-ban-${t.id}.png`}
                                        className="w-7 h-7 bg-slate-100 text-slate-500 rounded-lg flex items-center justify-center hover:bg-slate-900 hover:text-white transition-all shadow-sm"
                                        title="Tải ảnh về"
                                      >
                                        <Download className="w-4 h-4" />
                                      </a>
                                    </div>
                                  )}
                                </div>
                              </td>
                            </tr>
                          )) : (
                            <tr>
                              <td colSpan={5} className="py-20 text-center text-slate-400 text-sm font-bold uppercase tracking-widest opacity-30">Không tìm thấy giao dịch phù hợp.</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </Card>
                </div>
              )}

              {activeTab === 'partners' && (
                <div className="space-y-6">
                  {isOwner && partners.length <= 1 && (
                    <div 
                      className="bg-amber-50 border border-amber-200 p-6 rounded-3xl flex flex-col sm:flex-row items-center justify-between gap-6"
                    >
                      <div className="flex items-center gap-4">
                        <div className="w-12 h-12 bg-amber-500 rounded-2xl flex items-center justify-center text-white shadow-lg">
                          <Users className="w-6 h-6" />
                        </div>
                        <div className="text-center sm:text-left">
                          <h3 className="text-sm font-black text-amber-900 uppercase tracking-widest">Danh sách Đối tác đang trống</h3>
                          <p className="text-xs font-bold text-amber-700/70 mt-1">Anh có muốn Tin khôi phục lại danh sách đối tác mẫu (SKB, APC, BNG...) không?</p>
                        </div>
                      </div>
                      <button 
                        onClick={handleRestorePartners}
                        className="px-8 py-3.5 bg-amber-600 text-white rounded-2xl text-[10px] font-black uppercase tracking-[0.2em] shadow-xl shadow-amber-600/20 hover:bg-amber-700 active:scale-95 transition-all flex items-center gap-2"
                      >
                        <RefreshCw className="w-4 h-4" /> Khôi phục ngay
                      </button>
                    </div>
                  )}

                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  {partners.map(p => (
                    <Card key={p.id} className="relative group">
                      <div className="flex items-start justify-between">
                        <div className="space-y-4">
                          <div className={cn(
                            "px-2 py-0.5 rounded text-[10px] font-black uppercase tracking-tighter inline-block",
                            p.type === 'SUPPLIER' ? "bg-blue-100 text-blue-600" :
                            p.type === 'RESTAURANT' ? "bg-amber-100 text-amber-600" :
                            p.type === 'AGENT' ? "bg-indigo-100 text-indigo-600" : "bg-gray-100 text-gray-600" 
                          )}>
                            {p.type === 'SUPPLIER' ? 'Nhà cung cấp' : 
                             p.type === 'RESTAURANT' ? 'Nhà hàng' :
                             p.type === 'AGENT' ? 'Đại lý' : 'Khách lẻ'}
                          </div>
                          <div className="flex items-center gap-2">
                            <h4 className="text-lg font-bold text-gray-900 leading-tight">{p.name}</h4>
                            {p.sapCode && <span className="px-1.5 py-0.5 bg-slate-100 text-slate-500 rounded text-[9px] font-black border border-slate-200">SAP: {p.sapCode}</span>}
                            <CheckCircle className="w-4 h-4 text-emerald-500 fill-emerald-50" />
                          </div>
                          <div>
                            <p className="text-sm text-gray-500 mt-1">{p.phone || 'Không có SĐT'}</p>
                          </div>
                          {p.address && (
                            <div className="flex items-start gap-2 text-xs text-gray-400">
                              <Search className="w-3 h-3 mt-0.5 shrink-0" />
                              <span>{p.address}</span>
                            </div>
                          )}
                        </div>
                        <button 
                          onClick={() => handleDeletePartner(p.id)}
                          className="p-2 hover:bg-rose-50 rounded-lg transition-colors group/del"
                          title="Xóa đối tác"
                        >
                          <Trash2 className="w-4 h-4 text-gray-300 group-hover/del:text-rose-500" />
                        </button>
                      </div>
                      <div className="mt-6 pt-4 border-t border-gray-50 flex items-center justify-between">
                         <div className="text-[10px] uppercase font-bold text-gray-400">Tổng GD</div>
                         <div className="text-sm font-bold text-gray-900">
                           {transactions.filter(t => t.partnerId === p.id).length}
                         </div>
                      </div>
                    </Card>
                  ))}
                  <button 
                    onClick={() => setShowAddPartner(true)}
                    className="border-2 border-dashed border-gray-200 rounded-xl p-8 flex flex-col items-center justify-center gap-3 text-gray-400 hover:border-blue-400 hover:text-blue-500 hover:bg-blue-50 transition-all group"
                  >
                    <PlusCircle className="w-8 h-8 group-hover:scale-110 transition-transform" />
                    <span className="text-sm font-medium">Thêm đối tác mới</span>
                  </button>
                </div>
              </div>
            )}

              {activeTab === 'gallery' && (
                <div className="space-y-8">
                  <div className="flex flex-col md:flex-row md:items-end justify-between gap-6">
                    <div>
                      <h2 className="text-3xl font-black text-slate-900 tracking-tight italic font-serif uppercase">
                        THƯ VIỆN ẢNH {galleryFilter === 'IN' ? 'NHẬP KHO' : 'XUẤT KHO'}
                      </h2>
                      <p className="text-[10px] font-black text-slate-400 mt-2 uppercase tracking-[0.3em]">
                        Minh chứng giao dịch thực hiện bởi BP {galleryFilter === 'IN' ? 'Nhập kho' : 'Xuất kho'}
                      </p>
                    </div>
                    
                    <div className="flex flex-col lg:flex-row items-center gap-4 w-full md:w-auto">
                      {/* Search Dynamic Field */}
                      <div className="relative group w-full lg:w-64">
                        <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 group-focus-within:text-primary transition-colors" />
                        <input 
                          type="text" 
                          placeholder={galleryFilter === 'IN' ? "Tra cứu MÃ LÔ..." : "Tra cứu ĐƠN VỊ..."} 
                          className="w-full pl-10 pr-4 py-2.5 bg-white border border-slate-200 rounded-xl text-[11px] sm:text-xs font-black placeholder:text-slate-300 focus:outline-none focus:ring-4 focus:ring-primary/5 focus:border-primary transition-all shadow-sm uppercase tracking-widest text-slate-700"
                          value={gallerySearchQuery}
                          onChange={(e) => setGallerySearchQuery(e.target.value)}
                        />
                        {gallerySearchQuery && (
                          <button 
                            onClick={() => setGallerySearchQuery('')}
                            className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-300 hover:text-slate-500"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        )}
                      </div>

                      {/* Month Filter */}
                      <div className="relative w-full lg:w-48">
                        <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 pointer-events-none" />
                        <select 
                          className="w-full pl-10 pr-4 py-2.5 bg-white border border-slate-200 rounded-xl text-[11px] sm:text-xs font-black focus:outline-none focus:ring-4 focus:ring-primary/5 focus:border-primary transition-all appearance-none cursor-pointer shadow-sm uppercase tracking-widest text-slate-700"
                          value={galleryMonth}
                          onChange={(e) => setGalleryMonth(e.target.value)}
                        >
                          {galleryMonthOptions.map(opt => (
                            <option key={opt.value} value={opt.value}>{opt.label.toUpperCase()}</option>
                          ))}
                        </select>
                        <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 pointer-events-none" />
                      </div>

                      {/* Toggle IN/OUT */}
                      <div className="flex bg-slate-100/50 backdrop-blur-sm p-1.5 rounded-[20px] sm:rounded-[24px] w-full md:w-auto border border-slate-200/50 shadow-inner">
                        <button 
                          onClick={() => setGalleryFilter('IN')}
                          className={cn(
                            "flex-1 md:px-8 py-3 rounded-lg sm:rounded-xl text-[10px] font-black uppercase tracking-widest transition-all duration-300",
                            galleryFilter === 'IN' 
                              ? "bg-white text-amber-600 shadow-xl shadow-amber-500/10 border border-slate-100" 
                              : "text-slate-400 hover:text-slate-600"
                          )}
                        >
                          Nhập kho
                        </button>
                        <button 
                          onClick={() => setGalleryFilter('OUT')}
                          className={cn(
                            "flex-1 md:px-8 py-3 rounded-lg sm:rounded-xl text-[10px] font-black uppercase tracking-widest transition-all duration-300",
                            galleryFilter === 'OUT' 
                              ? "bg-white text-rose-600 shadow-xl shadow-rose-500/10 border border-slate-100" 
                              : "text-slate-400 hover:text-slate-600"
                          )}
                        >
                          Xuất kho
                        </button>
                      </div>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4 sm:gap-6">
                    {filteredGalleryTransactions.length > 0 ? (
                      filteredGalleryTransactions.map((t) => (
                        <div 
                          key={t.id}
                          className="group relative bg-white rounded-3xl overflow-hidden shadow-sm hover:shadow-2xl transition-all border border-slate-100 cursor-pointer"
                          onClick={() => setSelectedGalleryImage(t)}
                        >
                          <div className="aspect-[4/5] overflow-hidden">
                            <img 
                              src={t.evidencePhotoUrl} 
                              alt={t.productName} 
                              className="w-full h-full object-cover transition-transform duration-700 group-hover:scale-110"
                            />
                            <div className="absolute inset-0 bg-gradient-to-t from-slate-900/80 via-transparent to-transparent opacity-60 group-hover:opacity-100 transition-opacity" />
                          </div>
                          
                          <div className="absolute bottom-0 left-0 right-0 p-4 sm:p-5">
                            <div className="flex flex-col gap-1">
                              <div className="flex items-center gap-1.5 mb-1">
                                <Calendar className="w-2.5 h-2.5 text-white/70" />
                                <span className={cn(
                                  "text-[9px] font-black uppercase tracking-widest",
                                  galleryFilter === 'IN' ? "text-amber-300" : "text-rose-300"
                                )}>{galleryFilter === 'IN' ? 'Ngày thực nhập' : 'Ngày thực xuất'}: {formatDate(t.date)}</span>
                              </div>
                              <h5 className="text-white font-bold text-[11px] sm:text-xs leading-tight truncate">{t.productName}</h5>
                              <div className="flex items-center gap-1.5 mt-1 border-t border-white/10 pt-2">
                                <Users className={cn(
                                  "w-2.5 h-2.5",
                                  galleryFilter === 'IN' ? "text-amber-400" : "text-rose-400"
                                )} />
                                <span className={cn(
                                  "text-[9px] font-black uppercase tracking-tighter truncate",
                                  galleryFilter === 'IN' ? "text-amber-200" : "text-rose-200"
                                )}>{t.partnerName}</span>
                              </div>
                            </div>
                          </div>

                          <div className="absolute top-4 right-4 opacity-0 group-hover:opacity-100 transition-all scale-75 group-hover:scale-100">
                             <div className="w-10 h-10 bg-white/20 backdrop-blur-md rounded-xl flex items-center justify-center text-white border border-white/20">
                               <ImageIcon className="w-5 h-5" />
                             </div>
                          </div>
                        </div>
                      ))
                    ) : (
                      <div className="col-span-full py-32 flex flex-col items-center justify-center text-center space-y-4 border-dashed border-2 bg-slate-50/50 rounded-[32px]">
                         <div className="w-20 h-20 bg-white rounded-3xl flex items-center justify-center shadow-xl text-slate-300">
                            <ImageIcon className="w-10 h-10" />
                         </div>
                         <div className="space-y-1">
                           <h4 className="text-xl font-black text-slate-900 uppercase">Trống</h4>
                           <p className="text-xs text-slate-400 font-bold uppercase tracking-widest">
                             Chưa có ảnh minh chứng {galleryFilter === 'IN' ? 'nhập kho' : 'xuất kho'} nào được ghi nhận.
                           </p>
                         </div>
                      </div>
                    )}
                  </div>
                </div>
              )}

              {/* Fullscreen Image View Modal */}
              {selectedGalleryImage && (
                <div className="fixed inset-0 z-[60] flex items-center justify-center p-4">
                  <div 
                    onClick={() => setSelectedGalleryImage(null)}
                    className="absolute inset-0 bg-slate-950/95 backdrop-blur-2xl"
                  />
                  <div 
                    className="relative w-full max-w-4xl max-h-[85vh] flex flex-col"
                  >
                    <div className="absolute -top-16 left-0 right-0 flex items-center justify-between pointer-events-none">
                       <div className="flex flex-col">
                          <h3 className="text-white font-black text-xl uppercase tracking-tighter">{selectedGalleryImage.productName}</h3>
                          <p className="text-white/40 text-[10px] font-black uppercase tracking-widest mt-1">
                            Ngày up: {formatDate(selectedGalleryImage.date)} • {selectedGalleryImage.partnerName}
                          </p>
                       </div>
                       <div className="flex gap-4 pointer-events-auto">
                          <a 
                            href={selectedGalleryImage.evidencePhotoUrl} 
                            download={`BCSX-${selectedGalleryImage.id}.png`}
                            className="w-12 h-12 bg-white/10 hover:bg-white text-white hover:text-slate-900 rounded-2xl flex items-center justify-center backdrop-blur-xl border border-white/10 transition-all"
                          >
                            <Download className="w-6 h-6" />
                          </a>
                          <button 
                            onClick={() => setSelectedGalleryImage(null)}
                            className="w-12 h-12 bg-rose-500 text-white rounded-2xl flex items-center justify-center shadow-2xl transition-all active:scale-90"
                          >
                            <X className="w-6 h-6" />
                          </button>
                       </div>
                    </div>
                    <img 
                      src={selectedGalleryImage.evidencePhotoUrl} 
                      className="w-full h-full object-contain rounded-3xl shadow-2xl"
                      alt="Zoomed" 
                    />
                  </div>
                </div>
              )}

              {/* Add Partner Modal */}
              {showAddPartner && (
                <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
                  <div 
                    onClick={() => setShowAddPartner(false)}
                    className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm"
                  />
                  <div 
                    className="relative w-full max-w-lg bg-white rounded-[32px] shadow-2xl overflow-hidden p-8"
                  >
                    <div className="flex items-center justify-between mb-8">
                      <div>
                        <h3 className="text-xl font-black text-slate-900 underline decoration-indigo-200 underline-offset-4">ĐỐI TÁC MỚI</h3>
                        <p className="text-[10px] font-black text-slate-400 mt-1 uppercase tracking-widest">Khởi rạo quan hệ hợp tác</p>
                      </div>
                      <button onClick={() => setShowAddPartner(false)} className="p-2 hover:bg-slate-100 rounded-full transition-colors">
                        <X className="w-5 h-5 text-slate-400" />
                      </button>
                    </div>

                    <form onSubmit={handleAddPartner} className="space-y-6">
                        <div className="grid grid-cols-3 gap-4">
                          <div className="col-span-2">
                            <Input 
                              label="Tên Đối tác / Đại lý"
                              placeholder="Nhập tên chính thức..."
                              value={partnerFormData.name}
                              onChange={(e: any) => setPartnerFormData({ ...partnerFormData, name: e.target.value })}
                              required
                            />
                          </div>
                          <Input 
                            label="Mã SAP (Tùy chọn)"
                            placeholder="AD0xxx..."
                            value={partnerFormData.sapCode}
                            onChange={(e: any) => setPartnerFormData({ ...partnerFormData, sapCode: e.target.value })}
                          />
                        </div>
                        <div className="grid grid-cols-2 gap-4">
                          <Select 
                            label="Phân loại"
                            options={[
                              { value: 'AGENT', label: 'Đại lý' },
                              { value: 'RESTAURANT', label: 'Nhà hàng' },
                              { value: 'SUPPLIER', label: 'Nhà cung cấp' },
                              { value: 'OTHER', label: 'Khác' }
                            ]}
                            value={partnerFormData.type}
                            onChange={(e: any) => setPartnerFormData({ ...partnerFormData, type: e.target.value as any })}
                          />
                          <Input 
                            label="Số điện thoại"
                            placeholder="09xx..."
                            value={partnerFormData.phone}
                            onChange={(e: any) => setPartnerFormData({ ...partnerFormData, phone: e.target.value })}
                          />
                        </div>
                        <Input 
                          label="Địa chỉ (Tùy chọn)"
                          placeholder="Nhập địa chỉ..."
                          value={partnerFormData.address}
                          onChange={(e: any) => setPartnerFormData({ ...partnerFormData, address: e.target.value })}
                        />
                        <div className="pt-4 flex gap-3">
                          <Button variant="outline" className="flex-1" type="button" onClick={() => setShowAddPartner(false)}>Hủy</Button>
                          <Button className="flex-[2]" type="submit">Lưu Đối tác</Button>
                        </div>
                      </form>
                    </div>
                  </div>
                )}
            </div>
        </div>
      </main>
    </div>
  );
}
