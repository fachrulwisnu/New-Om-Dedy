import React, { useState, useMemo, useEffect, Component, ErrorInfo } from 'react';
import * as XLSX from 'xlsx';

const handleExcelExport = (data: any[], fileName: string) => {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data");
  XLSX.writeFile(wb, `${fileName}.xlsx`);
};

import { HOLIDAYS_2026, CUTI_BERSAMA_2026 } from './constants';
import { 
  Plus, 
  History,
  LayoutGrid,
  AlertTriangle,
  Cpu,
  LayoutDashboard,
  Users,
  ShieldAlert,
  FolderKanban,
  ChevronLeft,
  ChevronRight,
  LogOut,
  Settings as SettingsIcon,
  Search,
  Trash2,
  ShieldCheck,
  ExternalLink,
  Edit3,
  Save,
  UserPlus,
  X,
  ArrowRight,
  User as UserIcon,
  Clock,
  ArrowLeft,
  ArrowDown,
  Rocket,
  Activity,
  Calendar,
  Layers,
  PlusSquare,
  Upload,
  Download,
  Filter,
  Copy
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  format, 
  addHours, 
  startOfDay, 
  endOfDay, 
  eachHourOfInterval, 
  isSameHour,
  differenceInHours,
  addDays,
  startOfWeek,
  endOfWeek,
  eachDayOfInterval,
  isSameDay,
  differenceInDays,
  startOfMonth,
  endOfMonth,
  eachWeekOfInterval,
  isWithinInterval,
  addMonths,
  isToday,
  isWeekend
} from 'date-fns';
import { Task, ViewScale, TaskStatus, ProjectStatus, Project, AppUser, AppView, AuditLog, Schedule, ProjectRescheduleLog, RescheduleRequest } from './types';
import { cn } from './lib/utils';
import { useTasks } from './hooks/useTasks';
import { taskService } from './services/taskService';
import { supabase } from './lib/supabase';

import { useAuth } from './context/AuthContext';
import { LoginPage } from './components/LoginPage';

const formatWorkday = (totalHours: number) => {
  const days = Math.floor(totalHours / 9);
  const remainingHours = Math.floor(totalHours % 9);
  const remainingMinutes = Math.round((totalHours % 1) * 60);
  
  const parts = [];
  if (days > 0) parts.push(`${days} Hari Kerja`);
  if (remainingHours > 0) parts.push(`${remainingHours} Jam`);
  if (remainingMinutes > 0) parts.push(`${remainingMinutes} Menit`);
  
  return parts.length > 0 ? parts.join(' ') : '0 Jam';
};

// --- Collision Engine ---
const getCollision = (currentTask: Task, allTasks: Task[], projects: Project[]) => {
  const overlap = allTasks.filter(task => {
    if (task.id === currentTask.id) return false;
    
    const start1 = new Date(currentTask.start_time);
    const end1 = new Date(currentTask.end_time);
    const start2 = new Date(task.start_time);
    const end2 = new Date(task.end_time);

    const timeOverlap = start1 < end2 && end1 > start2;
    if (!timeOverlap) return false;

    const samePIC = currentTask.assignee && task.assignee && currentTask.assignee === task.assignee;
    const sameDev = currentTask.developer_name && task.developer_name && currentTask.developer_name === task.developer_name;
    const sameQA = currentTask.qa_name && task.qa_name && currentTask.qa_name === task.qa_name;

    return samePIC || sameDev || sameQA;
  });

  if (overlap.length === 0) return null;

  return overlap.map(collisionTask => {
    const project = projects.find(p => p.id === collisionTask.project_id);
    return {
      ...collisionTask,
      projectName: project?.name || 'Unknown Project'
    };
  });
};

// --- Components ---

// --- Error Boundary ---
interface ErrorBoundaryProps {
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
}

class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false };
  }

  static getDerivedStateFromError(_: Error): ErrorBoundaryState {
    return { hasError: true };
  }

  componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error("ErrorBoundary caught an error:", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-slate-950/80 backdrop-blur-sm p-4">
          <div className="bg-slate-900 border border-slate-800 p-8 rounded-2xl max-w-md w-full shadow-2xl text-center space-y-6">
            <div className="w-16 h-16 bg-rose-500/10 rounded-full flex items-center justify-center mx-auto">
              <AlertTriangle className="w-8 h-8 text-rose-500" />
            </div>
            <div className="space-y-2">
              <h2 className="text-xl font-bold text-white tracking-tight">Wizard Critical Failure</h2>
              <p className="text-slate-400 text-sm leading-relaxed">
                The interface encountered a runtime error or missing asset. The component has been isolated for safety.
              </p>
            </div>
            <button 
              onClick={() => window.location.reload()}
              className="w-full py-3 bg-slate-800 hover:bg-slate-700 text-white font-bold uppercase text-xs tracking-widest rounded-xl transition-all"
            >
              Relational Reset (Reload)
            </button>
          </div>
        </div>
      );
    }

    return this.props.children; 
  }
}

const StatusBadge = ({ status, type = 'task' }: { status: string, type?: 'project' | 'task' }) => {
  const styles: Record<string, string> = {
    // Task Status
    'On Hold': 'bg-slate-500/10 text-slate-400 border-slate-500/20',
    'On Progress': 'bg-indigo-500/10 text-indigo-400 border-indigo-500/20',
    'Done': 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20',
    // Project Status
    'FSD on Progress': 'bg-blue-500/10 text-blue-400 border-blue-500/20',
    'FSD on Review': 'bg-purple-500/10 text-purple-400 border-purple-500/20',
    'SIT on Progress': 'bg-amber-500/10 text-amber-400 border-amber-500/20',
    'UAT on Progress': 'bg-indigo-500/10 text-indigo-400 border-indigo-500/20',
    'Project Late': 'bg-rose-500/10 text-rose-400 border-rose-500/20',
  };

  return (
    <span className={cn(
      "px-2 py-0.5 rounded-full text-[9px] font-black border uppercase tracking-wider whitespace-nowrap",
      styles[status] || 'bg-slate-500/10 text-slate-400 border-slate-500/20'
    )}>
      {status || 'Unknown'}
    </span>
  );
};

const TaskStatusSelector = ({ status, onUpdate }: { status: any, onUpdate: (s: any) => void }) => {
  const options = [TaskStatus.ON_HOLD, TaskStatus.ON_PROGRESS, TaskStatus.DONE];
  return (
    <select 
      value={status || TaskStatus.ON_PROGRESS}
      onChange={(e) => onUpdate(e.target.value)}
      className={cn(
        "bg-slate-900 border border-slate-800 rounded px-2 py-1 text-[9px] font-black uppercase outline-none transition-all cursor-pointer",
        status === TaskStatus.DONE ? "text-emerald-400" : status === TaskStatus.ON_HOLD ? "text-rose-400" : "text-indigo-400"
      )}
    >
      {options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
    </select>
  );
};

const ProjectStatusSelector = ({ status, onUpdate }: { status: any, onUpdate: (s: any) => void }) => {
  const options = [
    ProjectStatus.FSD_PROGRESS, 
    ProjectStatus.FSD_REVIEW, 
    ProjectStatus.SIT_PROGRESS, 
    ProjectStatus.ON_HOLD, 
    ProjectStatus.UAT_PROGRESS
  ];
  return (
    <select 
      value={status || ProjectStatus.FSD_PROGRESS}
      onChange={(e) => onUpdate(e.target.value)}
      className="bg-slate-900 border border-slate-800 rounded px-2 py-1 text-[9px] font-black uppercase text-indigo-400 outline-none transition-all cursor-pointer"
    >
      {options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
    </select>
  );
};

const ApprovalBadge = ({ value, label, onUpdate }: { value: string | null, label: string, onUpdate: (val: string) => void }) => {
  const options = ['Pending', 'Revise', 'OK'];
  const safeValue = options.includes(value || '') ? (value || 'Pending') : 'Pending';

  return (
    <div className="flex flex-col gap-1 w-full max-w-[80px]">
      <select 
        value={safeValue}
        onChange={(e) => onUpdate(e.target.value)}
        className={cn(
          "bg-slate-900 border border-slate-800 rounded px-1.5 py-1 text-[9px] font-black uppercase tracking-tighter outline-none focus:ring-1 focus:ring-indigo-500/50 transition-all cursor-pointer w-full",
          safeValue === 'OK' ? "text-emerald-400 border-emerald-500/30" :
          safeValue === 'Revise' ? "text-rose-400 border-rose-500/30" :
          "text-slate-500"
        )}
      >
        {options.map(opt => (
          <option key={opt} value={opt} className="bg-slate-900">{opt}</option>
        ))}
      </select>
    </div>
  );
};


// --- Main App ---

// --- Health Engine ---

const getTaskHealth = (task: Task) => {
  if (task.status === TaskStatus.DONE) return 'Stable';
  
  const now = new Date();
  const end = new Date(task.end_time);
  
  if (now > end) return 'OVERDUE';
  
  if (task.target_sla_date) {
    const sla = new Date(task.target_sla_date);
    if (now > sla) return 'OVER SLA';
  }
  
  return 'Healthy';
};

const HealthBadge = ({ health }: { health: string }) => {
  if (health === 'Stable') return null;
  if (health === 'Healthy') return null;
  
  return (
    <span className={cn(
      "px-2 py-0.5 rounded-full text-[9px] font-black uppercase tracking-widest animate-pulse shadow-lg",
      health === 'OVERDUE' ? "bg-rose-500 text-white shadow-rose-500/20" : "bg-amber-500 text-slate-900 shadow-amber-500/20"
    )}>
      {health}
    </span>
  );
};

// --- Custom UI Overlays ---

const ConfirmModal = ({ 
  isOpen, 
  onClose, 
  onConfirm, 
  title, 
  description, 
  confirmText = 'Hapus', 
  cancelText = 'Batal',
  variant = 'danger'
}: { 
  isOpen: boolean, 
  onClose: () => void, 
  onConfirm: () => void, 
  title: string, 
  description: string,
  confirmText?: string,
  cancelText?: string,
  variant?: 'danger' | 'primary'
}) => {
  if (!isOpen) return null;

  return (
    <AnimatePresence>
      <div className="fixed inset-0 z-[150] flex items-center justify-center p-4">
        <motion.div 
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          onClick={onClose}
          className="absolute inset-0 bg-slate-950/80 backdrop-blur-sm"
        />
        <motion.div 
          initial={{ opacity: 0, scale: 0.95, y: 10 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.95, y: 10 }}
          className="relative w-full max-w-md bg-slate-900 border border-slate-700/50 rounded-xl shadow-2xl overflow-hidden p-6"
        >
          <div className="flex flex-col items-center text-center">
            <div className={cn(
              "w-12 h-12 rounded-full flex items-center justify-center mb-4",
              variant === 'primary' ? "bg-indigo-500/10" : "bg-red-500/10"
            )}>
              {variant === 'primary' ? (
                <ShieldCheck className="w-6 h-6 text-indigo-500" />
              ) : (
                <AlertTriangle className="w-6 h-6 text-red-500" />
              )}
            </div>
            <h3 className="text-xl font-bold text-slate-100 mb-2">{title}</h3>
            <p className="text-slate-400 mb-6 text-sm leading-relaxed">{description}</p>
          </div>
          
          <div className="flex justify-end gap-3">
            <button 
              onClick={onClose}
              className="bg-slate-800 hover:bg-slate-700 text-slate-200 px-4 py-2 rounded-lg transition-colors text-sm font-semibold"
            >
              {cancelText}
            </button>
            <button 
              onClick={() => { onConfirm(); onClose(); }}
              className={cn(
                "px-4 py-2 rounded-lg transition-colors flex items-center gap-2 text-sm font-semibold",
                variant === 'primary' 
                  ? "bg-indigo-600 hover:bg-indigo-500 text-white" 
                  : "bg-red-500/10 hover:bg-red-600 border border-red-500/20 text-red-500 hover:text-white"
              )}
            >
              {variant === 'danger' && <Trash2 className="w-4 h-4" />}
              {confirmText}
            </button>
          </div>
        </motion.div>
      </div>
    </AnimatePresence>
  );
};

const SuccessNotification = ({ show, message, onClose }: { show: boolean, message: string, onClose: () => void }) => {
  useEffect(() => {
    if (show) {
      const timer = setTimeout(onClose, 3000);
      return () => clearTimeout(timer);
    }
  }, [show, onClose]);

  return (
    <AnimatePresence>
      {show && (
        <motion.div 
          initial={{ opacity: 0, y: -50, x: '-50%' }}
          animate={{ opacity: 1, y: 0, x: '-50%' }}
          exit={{ opacity: 0, y: -50, x: '-50%' }}
          className="fixed top-8 left-1/2 z-[150] bg-emerald-500 text-white px-6 py-3 rounded-full shadow-[0_10px_40px_rgba(16,185,129,0.4)] flex items-center gap-3 border border-emerald-400"
        >
          <div className="w-5 h-5 bg-white/20 rounded-full flex items-center justify-center">
            <Plus className="w-3 h-3 rotate-45" />
          </div>
          <span className="text-xs font-black uppercase tracking-widest">{message}</span>
        </motion.div>
      )}
    </AnimatePresence>
  );
};

// --- Components ---

function EditableInput({ value, onSave, className, placeholder, type = 'text', min, max }: { 
  value: string | number, 
  onSave: (val: any) => void, 
  className?: string, 
  placeholder?: string,
  type?: 'text' | 'number',
  min?: string | number,
  max?: string | number
}) {
  const [localValue, setLocalValue] = useState(value);

  useEffect(() => {
    setLocalValue(value);
  }, [value]);

  const handleBlur = () => {
    if (localValue !== value) {
      if (type === 'number') {
        let val = parseFloat(String(localValue)) || 0;
        if (min !== undefined) val = Math.max(Number(min), val);
        if (max !== undefined) val = Math.min(Number(max), val);
        onSave(val);
      } else {
        onSave(localValue);
      }
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      (e.target as HTMLInputElement).blur();
    }
    if (e.key === 'Escape') {
      setLocalValue(value);
      (e.target as HTMLInputElement).blur();
    }
  };

  return (
    <input 
      type={type}
      min={min}
      value={localValue}
      onChange={(e) => {
        const val = e.target.value;
        if (type === 'number' && min !== undefined) {
           const n = parseInt(val) || 0;
           setLocalValue(Math.max(Number(min), n));
        } else {
           setLocalValue(val);
        }
      }}
      onBlur={handleBlur}
      onKeyDown={handleKeyDown}
      className={className}
      placeholder={placeholder}
    />
  );
}

const CollisionWarning = ({ collisions }: { collisions: any[] }) => {
  return (
    <div className="relative group inline-flex items-center justify-center ml-2">
      <button 
        className="text-rose-500 animate-pulse hover:scale-110 transition-transform"
      >
        <AlertTriangle className="w-4 h-4" />
      </button>
      
      <div className="absolute bottom-full mb-2 hidden group-hover:block w-max max-w-xs bg-slate-800 border border-slate-700 text-slate-100 text-xs rounded-md p-3 z-[100] shadow-xl pointer-events-none">
        <p className="text-[10px] font-bold text-rose-400 uppercase tracking-widest mb-2 flex items-center gap-1">
          <ShieldAlert className="w-3 h-3" /> External Conflict Detected
        </p>
        <div className="space-y-3">
          {(collisions || []).map((c, i) => {
            if (!c) return null;
            return (
              <div key={`${c.id}-${i}`} className="text-[11px] leading-relaxed border-l-2 border-slate-800 pl-2">
                <span className="text-slate-300 font-bold block mb-0.5">Project: {c.projectName || 'Unknown'}</span>
                <span className="text-slate-500 block">Task: <span className="text-indigo-400">"{c.title || 'Untitled'}"</span></span>
                <span className="text-slate-500 block">Conflict: {c.start_time ? format(new Date(c.start_time), 'MMM dd') : '??'} - {c.end_time ? format(new Date(c.end_time), 'MMM dd') : '??'}</span>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};

export default function App() {
  const { user: authUser, loading: authLoading, signOut: realSignOut } = useAuth();
  
  // MOCK USER for temporary setup bypass
  const user = useMemo(() => authUser || (({
    id: 'mock-admin-id',
    name: 'Fachrul Wisnu Novianto',
    email: 'fachrulwisnunovianto@gmail.com',
    access_level: 'Superadmin',
    role: 'Product Manager'
  }) as any), [authUser]);

  const signOut = () => {
    if (authUser) realSignOut();
    else {
      console.log("Mock session logout requested");
      // Optionally reset some local state if needed, but for bypass we stay logged in
    }
  };
  const [activeView, setActiveView] = useState<AppView>('PROJECTS');
  const [selectedProjectId, setSelectedProjectId] = useState<string | null>(null);
  const [focusedProjectId, setFocusedProjectId] = useState<string | null>(null);
  const [notif, setNotif] = useState<string | null>(null);
  
  const [projects, setProjects] = useState<Project[]>([]);
  const [users, setUsers] = useState<AppUser[]>([]);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [allAuditLogs, setAllAuditLogs] = useState<AuditLog[]>([]);

  // Automatic "Project Late" Detection
  useEffect(() => {
    const checkLateProjects = async () => {
      const lateProjects = projects.filter(p => {
        if (p.status === 'Project Late') return false;
        const projectTasks = tasks.filter(t => t.project_id === p.id);
        if (projectTasks.length === 0) return false;
        
        // Find the latest task end date
        const latestEnd = Math.max(...projectTasks.map(t => new Date(t.end_time).getTime()));
        return latestEnd < new Date().getTime();
      });

      for (const p of lateProjects) {
        await taskService.updateProject(p.id, { status: ProjectStatus.PROJECT_LATE }, 'System Monitor');
      }
      if (lateProjects.length > 0) setRefreshKey(prev => prev + 1);
    };

    if (projects.length > 0 && tasks.length > 0) {
      checkLateProjects();
    }
  }, [projects, tasks]);

  
  const [showAuditLog, setShowAuditLog] = useState(false);
  const [selectedTask, setSelectedTask] = useState<Task | null>(null);
  const [auditLogs, setAuditLogs] = useState<AuditLog[]>([]);
  
  const [loading, setLoading] = useState(true);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  
  const [scale, setScale] = useState<ViewScale>('DAY');
  const [expandedRows, setExpandedRows] = useState<Set<string>>(new Set());
  const [refreshKey, setRefreshKey] = useState(0);
  const [rescheduleRequests, setRescheduleRequests] = useState<any[]>([]);

  const fetchRescheduleRequests = async () => {
    if (!canAccessReschedule) return;
    try {
      const data = await taskService.getRescheduleRequests();
      setRescheduleRequests(data);
    } catch (err) {
      console.error("Failed to fetch reschedule requests:", err);
    }
  };
  const [isCreateProjectModalOpen, setIsCreateProjectModalOpen] = useState(false);
  const [taskModalData, setTaskModalData] = useState<{ parentId: string | null } | null>(null);
  const [reschedulingProject, setReschedulingProject] = useState<Project | null>(null);
  const [itemToDelete, setItemToDelete] = useState<{id: string, type: 'task' | 'project' | 'user' | 'phase' | 'subtask', phaseIdx?: number, subIdx?: number} | null>(null);

  const requestDeleteTask = (id: string) => {
    setItemToDelete({ id, type: 'task' });
  };

  const confirmDelete = () => {
    if (!itemToDelete) return;
    const { id, type, phaseIdx, subIdx } = itemToDelete;
    
    if (type === 'task') handleDeleteTask(id);
    else if (type === 'project') handleDeleteProject(id);
    else if (type === 'user') handleDeleteUser(id);
    
    setItemToDelete(null);
  };

  // Auth guard temporarily disabled for setup purposes
  // if (authLoading) return (
  //   <div className="min-h-screen bg-[#020617] flex items-center justify-center">
  //     <div className="w-8 h-8 border-4 border-indigo-500/20 border-t-indigo-500 rounded-full animate-spin" />
  //   </div>
  // );

  // if (!user) return <LoginPage />;

  const fetchData = async () => {
    try {
      const [p, u, t, a] = await Promise.all([
        taskService.getProjects(),
        taskService.getUsers(),
        taskService.getAllTasks(),
        taskService.getAuditLogs()
      ]);
      setProjects(p);
      setUsers(u);
      setTasks(t);
      setAllAuditLogs(a);
    } catch (err) {
      console.error("Fetch failed:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();

    const channel = supabase
      .channel('portfolio_realtime')
      .on('postgres_changes', { event: '*', schema: 'public' }, () => setRefreshKey(prev => prev + 1))
      .subscribe();

    const requestChannel = supabase.channel('reschedule_realtime')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'reschedule_requests' }, () => fetchRescheduleRequests())
      .subscribe();

    fetchRescheduleRequests();

    return () => {
      supabase.removeChannel(channel);
      supabase.removeChannel(requestChannel);
    };
  }, [refreshKey]);

  // Bootstrap Superadmin "fachrul wisnu"
  useEffect(() => {
    const bootstrap = async () => {
      if (loading || users.length === 0) return;
      const targetEmail = 'fachrulwisnunovianto@gmail.com';
      const exists = users.some(u => u.email?.toLowerCase() === targetEmail.toLowerCase());
      
      if (!exists) {
        console.log("Bootstrapping Superadmin: fachrul wisnu");
        try {
          await taskService.createUser({
            name: 'fachrul wisnu',
            email: targetEmail,
            access_level: 'Superadmin',
            role: 'Product Manager',
            password: 'bosskubabi'
          }, 'System Bootstrap');
          setRefreshKey(prev => prev + 1);
        } catch (err) {
          console.error("Bootstrap failed:", err);
        }
      }
    };
    bootstrap();
  }, [loading, users.length]);

  const currentProject = useMemo(() => 
    (projects || []).find(p => p.id === selectedProjectId), 
  [projects, selectedProjectId]);

  const currentUserProfile = useMemo(() => {
    if (!user || (users || []).length === 0) return null;
    return (users || []).find(u => u.email?.toLowerCase() === user.email?.toLowerCase());
  }, [user, users]);

  const userAccess = useMemo(() => 
    (currentUserProfile?.access_level || user?.access_level || '').toLowerCase().trim(),
  [currentUserProfile, user]);

  const isAdmin = useMemo(() => {
    return userAccess === 'superadmin' || userAccess === 'admin' || user?.email?.toLowerCase().includes('wisnu');
  }, [userAccess, user]);

  const canAccessReschedule = useMemo(() => 
    userAccess === 'admin' || userAccess === 'superadmin',
  [userAccess]);

  const filteredTasks = useMemo(() => 
    (tasks || []).filter(t => t.project_id === selectedProjectId),
  [tasks, selectedProjectId]);

  const hierarchicalTasks = useMemo(() => {
    const map = new Map<string, Task[]>();
    const roots: Task[] = [];
    
    (filteredTasks || []).forEach(t => {
      if (t.parent_id) {
        const children = map.get(t.parent_id) || [];
        children.push(t);
        map.set(t.parent_id, children);
      } else {
        roots.push(t);
      }
    });

    return { roots, map };
  }, [filteredTasks]);

  const handleUpdateTask = async (id: string, field: keyof Task, value: any) => {
    // Get existing task for audit/date check
    const task = tasks.find(t => t.id === id);
    const oldValue = task ? task[field] : null;

    // Optimistic Update
    setTasks(prev => prev.map(t => t.id === id ? { ...t, [field]: value } : t));
    
    try {
       await taskService.updateTask(id, { [field]: value }, user?.email || 'System User');
       
       // Log to Project Reschedule Log if date changed
       if (task && (field === 'start_time' || field === 'end_time')) {
         if (oldValue !== value && task.project_id) {
           await taskService.createProjectRescheduleLog({
             project_id: task.project_id,
             changed_by: user?.email || user?.name || 'User',
             old_start_date: field === 'start_time' ? format(new Date(String(oldValue)), 'yyyy-MM-dd') : format(new Date(task.start_time), 'yyyy-MM-dd'),
             old_end_date: field === 'end_time' ? format(new Date(String(oldValue)), 'yyyy-MM-dd') : format(new Date(task.end_time), 'yyyy-MM-dd'),
             new_start_date: field === 'start_time' ? format(new Date(String(value)), 'yyyy-MM-dd') : format(new Date(task.start_time), 'yyyy-MM-dd'),
             new_end_date: field === 'end_time' ? format(new Date(String(value)), 'yyyy-MM-dd') : format(new Date(task.end_time), 'yyyy-MM-dd'),
             reason: `Manual update on task: ${task.title}`
           });
         }
       }
    } catch (err: any) {
       console.error('Update failed:', err);
       fetchData(); 
       alert('CRUD Failed (Update Task): ' + err.message);
    }
  };

  const handleOpenAudit = async (task: Task) => {
    setSelectedTask(task);
    setShowAuditLog(true);
    try {
      const logs = await taskService.getAuditLogs({ taskId: task.id });
      setAuditLogs(logs);
    } catch (err) {
      console.error('Audit fetch failed:', err);
    }
  };

  const handleDeleteTask = async (id: string) => {
    // Optimistic Delete - disappear instantly from UI
    // Cascade removal: filter out the task itself AND any tasks that have it as parent_id
    setTasks(prevTasks => prevTasks.filter(t => t.id !== id && t.parent_id !== id));
    
    try {
      const { error } = await supabase.from('tasks').delete().eq('id', id);
      if (error) throw error;

      // Also log audit via service for traceability
      await taskService.logAudit({ 
        task_id: id, 
        actor: user?.email || 'Administrator', 
        action: 'DELETED' 
      });
      
    } catch (err: any) {
      console.error("Delete failed:", err);
      // Revert if failed to ensure data integrity
      fetchData(); 
      alert("CRUD Error: Gagal menghapus task. Database rejected the request.\nDetail: " + err.message);
    }
  };

  const handleAddUser = async () => {
    try {
      await taskService.createUser({ 
        name: 'New Personnel...', 
        access_level: 'PIC', 
        role: 'Staff' 
      }, 'User');
      setRefreshKey(prev => prev + 1);
    } catch (err: any) {
      console.error("Failed to add user:", err);
    }
  };

  const handleDeleteUser = async (id: string) => {
    try {
      await taskService.deleteUser(id, 'User');
      setRefreshKey(prev => prev + 1);
    } catch (err: any) {
      console.error("Failed to delete user:", err);
    }
  };

  const handleUpdateUser = async (id: string, field: keyof AppUser, value: any) => {
    try {
      await taskService.updateUser(id, { [field]: value }, 'User');
      setRefreshKey(prev => prev + 1);
    } catch (err: any) {
      console.error("Failed to update user:", err);
    }
  };

  const handleUpdateProject = async (id: string, updates: Partial<Project>) => {
    // Optimistic Update
    setProjects(prev => prev.map(p => p.id === id ? { ...p, ...updates } : p));
    
    try {
      await taskService.updateProject(id, updates, 'User');
      // refreshKey handled by subscriber
    } catch (err: any) {
      console.error("Failed to update project:", err);
      fetchData(); // Rollback
    }
  };

  const handleDeleteProject = async (id: string) => {
    // Optimistic Delete - disappear instantly
    setProjects(prev => prev.filter(p => p.id !== id));
    if (selectedProjectId === id) setSelectedProjectId(null);

    try {
      const { error } = await supabase.from('projects').delete().eq('id', id);
      if (error) throw error;
      setNotif('Project berhasil dihapus');
    } catch (err: any) {
      console.error("Failed to delete project:", err);
      setNotif('Gagal menghapus project');
      fetchData(); // Rollback
    }
  };

  const handleToggleExpand = (id: string) => {
    const next = new Set(expandedRows);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    setExpandedRows(next);
  };

  const menuItems = useMemo(() => {
    const base = [
      { id: 'PROJECTS', label: 'Project List', icon: LayoutDashboard },
      { id: 'GANTT_DETAIL', label: 'Timeline View', icon: Activity },
      { id: 'KANBAN', label: 'Status Monitoring', icon: LayoutGrid },
      { id: 'SCHEDULE', label: 'Om Dedy Schedule', icon: Calendar },
      { id: 'RESCHEDULE', label: 'Reschedule Om Dedy', icon: History },
      { id: 'PERSONEL', label: 'Personel OM DEDY', icon: Users },
      { id: 'AUDIT', label: 'System Audit Logs', icon: ShieldAlert },
    ];

    const filtered = base.filter(item => {
      // Only Admin/Superadmin can see Personnel and Reschedule menus
      if (item.id === 'PERSONEL' || item.id === 'RESCHEDULE' || item.id === 'AUDIT') return isAdmin;
      return true;
    });

    return filtered;
  }, [isAdmin]);

  if (loading && projects.length === 0) {
    return (
      <div className="fixed inset-0 bg-[#0B1120] flex flex-col items-center justify-center z-[9999]">
        <motion.div 
          initial={{ opacity: 0, scale: 0.9 }}
          animate={{ opacity: 1, scale: 1 }}
          className="flex flex-col items-center"
        >
          {/* Logo with Glow */}
          <div className="w-20 h-20 bg-indigo-600 rounded-2xl flex items-center justify-center mb-6 shadow-[0_0_40px_rgba(79,70,229,0.4)] relative overflow-hidden group">
             <span className="text-3xl font-black text-white relative z-10">OD</span>
             <div className="absolute inset-0 bg-gradient-to-tr from-white/20 to-transparent" />
          </div>

          {/* Branded Text */}
          <h1 className="text-3xl font-black tracking-tighter text-white mb-2">
            OM <span className="text-indigo-500">DEDY</span>
          </h1>
          <p className="text-slate-400 text-sm font-medium tracking-widest uppercase">
            Operational Monitoring Dashboard
          </p>
          <p className="text-slate-500 text-[10px] mt-1 font-bold tracking-[0.2em] uppercase">
            FOR EFFICIENT DELIVERY
          </p>

          {/* Modern Progress Bar */}
          <div className="w-48 h-1 bg-slate-800 rounded-full mt-10 overflow-hidden relative">
             <div 
               className="h-full bg-indigo-500 absolute top-0 left-0 animate-[loading_1.5s_infinite] w-1/3 shadow-[0_0_10px_#6366f1]" 
             />
          </div>
          
          <p className="text-[9px] text-slate-600 mt-6 font-mono uppercase tracking-widest animate-pulse">
            Synchronizing Critical Systems...
          </p>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#020617] text-slate-200 flex overflow-hidden">
      {/* Sidebar */}
      <aside className={cn(
        "bg-slate-900 border-r border-slate-800 transition-all duration-300 flex flex-col z-40",
        isSidebarOpen ? "w-72" : "w-16"
      )}>
        <div className="h-20 flex items-center px-4 border-b border-slate-800 shrink-0">
          <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center shrink-0 shadow-lg shadow-indigo-500/20 group cursor-pointer hover:scale-105 transition-transform">
            <span className="text-lg font-black text-white">OD</span>
          </div>
          {isSidebarOpen && (
            <div className="ml-3 overflow-hidden">
              <h1 className="font-black text-white uppercase italic tracking-tighter text-xl leading-none">OM <span className="text-indigo-500">DEDY</span></h1>
              <p className="text-[7px] text-slate-500 font-black tracking-[0.05em] uppercase leading-tight mt-1">Operational Monitoring Dashboard<br/>for Efficient Delivery</p>
            </div>
          )}
        </div>

        <nav className="flex-1 py-6 space-y-1">
          {menuItems.map((item, i) => (
            <button
              key={`${item.id}-${i}`}
              onClick={() => { setActiveView(item.id as AppView); setSelectedProjectId(null); }}
              className={cn(
                "w-full flex items-center py-3 px-4 transition-all relative group",
                activeView === item.id && !selectedProjectId 
                  ? "text-white bg-gradient-to-r from-indigo-600/20 to-transparent" 
                  : "text-slate-400 hover:text-slate-200 hover:bg-white/5"
              )}
            >
              <item.icon className={cn("w-5 h-5 shrink-0 transition-colors", activeView === item.id ? "text-indigo-400" : "text-slate-500")} />
              {isSidebarOpen && <span className="ml-4 font-bold text-xs uppercase tracking-widest">{item.label}</span>}
              {!isSidebarOpen && (
                <div className="absolute left-full ml-2 px-3 py-1 bg-slate-800 text-white text-[10px] rounded opacity-0 group-hover:opacity-100 whitespace-nowrap shadow-xl pointer-events-none z-50">
                  {item.label}
                </div>
              )}
              {activeView === item.id && !selectedProjectId && <div className="absolute left-0 top-0 bottom-0 w-1 bg-indigo-500" />}
            </button>
          ))}
        </nav>

        <div className="p-4 border-t border-slate-800 space-y-4">
           {isSidebarOpen && (
             <div className="bg-slate-950/50 p-4 rounded-xl border border-slate-800">
               <div className="flex items-center gap-3 mb-3">
                 <div className="w-8 h-8 rounded-full bg-slate-800 border border-indigo-500/30 flex items-center justify-center overflow-hidden">
                   <span className="text-[10px] font-black text-indigo-400 capitalize">{user.email?.charAt(0)}</span>
                 </div>
                 <div className="flex-1 min-w-0">
                   <p className="text-[10px] font-bold text-slate-200 truncate">{user.email}</p>
                   <p className="text-[9px] text-slate-500 uppercase tracking-widest font-bold">Authenticated PIC</p>
                 </div>
               </div>
               <button 
                 onClick={() => signOut()}
                 className="w-full flex items-center justify-center gap-2 py-2 bg-slate-800 hover:bg-rose-500/10 hover:text-rose-400 rounded-lg text-[10px] font-bold uppercase transition-all"
               >
                 <LogOut className="w-3.5 h-3.5" /> Log Out
               </button>
             </div>
           )}
           <button 
             onClick={() => setIsSidebarOpen(!isSidebarOpen)}
             className="w-full py-2 flex items-center justify-center rounded-lg hover:bg-slate-800 transition-colors"
           >
             {isSidebarOpen ? <ChevronLeft className="w-5 h-5 text-slate-600" /> : <ChevronRight className="w-5 h-5 text-slate-600" />}
           </button>
        </div>
      </aside>

      {/* Content Area */}
      <main className="flex-1 flex flex-col overflow-hidden relative">
        {loading && (
          <div className="absolute inset-0 bg-slate-950/60 backdrop-blur-[4px] z-[100] flex items-center justify-center">
            <div className="flex flex-col items-center gap-4">
              <div className="w-10 h-10 border-2 border-indigo-500 border-t-transparent rounded-full animate-spin shadow-[0_0_20px_rgba(99,102,241,0.3)]" />
              <div className="text-center">
                <span className="text-[10px] font-black uppercase tracking-[0.3em] text-indigo-500">Syncing Node</span>
                <p className="text-[9px] text-slate-500 font-bold uppercase tracking-widest mt-1">Operational Monitoring Active</p>
              </div>
            </div>
          </div>
        )}
        <header className="h-20 border-b border-slate-800/60 flex items-center justify-between px-8 bg-slate-950/50 backdrop-blur-md z-30 shrink-0">
          <div className="flex items-center gap-4">
             {activeView === 'GANTT_DETAIL' && focusedProjectId && (
               <button 
                 onClick={() => setFocusedProjectId(null)}
                 className="p-2 hover:bg-slate-800 rounded-lg transition-colors text-slate-500 hover:text-white"
               >
                 <ArrowLeft className="w-5 h-5" />
               </button>
             )}
             <div>
               <h2 className="text-xl font-black text-white tracking-tighter flex items-center gap-3 uppercase italic">
                 {activeView === 'PROJECTS' && <span>OM <span className="text-indigo-500">DEDY</span></span>}
                 {activeView === 'KANBAN' && 'Project Status Monitoring'}
                 {activeView === 'PERSONEL' && 'Personel OM DEDY'}
                 {activeView === 'RESCHEDULE' && 'Reschedule Om Dedy'}
                 {activeView === 'SCHEDULE' && 'Om Dedy Schedule'}
                 {activeView === 'AUDIT' && 'System Audit Rails'}
                 {activeView === 'GANTT_DETAIL' && (
                    focusedProjectId ? (
                      <div className="flex items-center gap-3">
                        <span>{projects.find(p => p.id === focusedProjectId)?.name}</span>
                        <div className="flex items-center gap-2 not-italic">
                          {(() => {
                            const p = projects.find(prj => prj.id === focusedProjectId);
                            const pTasks = tasks.filter(t => t.project_id === focusedProjectId);
                            const totalHours = pTasks.reduce((sum, t) => sum + (t.duration_hours || 0), 0);
                            return (
                              <>
                                <div className="bg-slate-800/80 px-2 py-1 rounded-md border border-slate-700 flex items-center gap-1.5 shadow-sm">
                                  <Clock className="w-3 h-3 text-indigo-400" />
                                  <span className="text-[10px] font-black text-slate-300 tracking-wider uppercase">
                                    ⏱️ {totalHours.toFixed(1)} HOURS
                                  </span>
                                </div>
                                <div className={cn(
                                  "px-2 py-1 rounded-md border text-[10px] font-black tracking-widest uppercase",
                                  p?.status === ProjectStatus.DONE ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-400" :
                                  p?.status === ProjectStatus.ACTIVE ? "bg-amber-500/10 border-amber-500/30 text-amber-400" :
                                  "bg-slate-800/50 border-slate-700 text-slate-500"
                                )}>
                                  {p?.status === ProjectStatus.ACTIVE ? 'FSD ON PROGRESS' : p?.status || 'UNKNOWN'}
                                </div>
                              </>
                            );
                          })()}
                        </div>
                      </div>
                    ) : 'Global Timeline'
                  )}
               </h2>
               <p className="text-[9px] text-slate-600 font-black uppercase tracking-[0.3em]">Operational Monitoring Dashboard for Efficient Delivery</p>
             </div>
          </div>

          <div className="flex items-center gap-4">
             {activeView === 'GANTT_DETAIL' && (
               <div className="flex bg-slate-900 p-1 rounded-lg border border-slate-800">
                 {(['MONTH', 'WEEK', 'DAY', 'HOUR'] as ViewScale[]).map((s, si) => (
                   <button
                     key={`header-scale-${s}-${si}`}
                     onClick={() => setScale(s)}
                     className={cn(
                       "px-3 py-1 rounded-md text-[10px] font-bold tracking-wider transition-all",
                       scale === s ? "bg-slate-800 text-indigo-400 shadow-sm" : "text-slate-500 hover:text-slate-300"
                     )}
                   >
                     {s}
                   </button>
                 ))}
               </div>
             )}
             
             {activeView === 'PROJECTS' && (
               <button 
                onClick={() => taskService.seedSampleData('System Architect')}
                className="bg-slate-800 hover:bg-slate-700 text-slate-200 px-4 py-2 rounded-lg text-xs font-bold transition-all border border-slate-700"
               >
                 Seed Infrastructure
               </button>
             )}
             
             <button className="p-2 text-slate-400 hover:text-white transition-colors">
               <SettingsIcon className="w-5 h-5" />
             </button>
          </div>
        </header>

        <section className="flex-1 overflow-auto p-8 scrollbar-hide">
          <ErrorBoundary>
           {activeView === 'PROJECTS' && (
             <PortfolioDashboard 
               projects={projects} 
               tasks={tasks}
               loading={loading}
               onOpenProject={(id) => { setSelectedProjectId(id); setFocusedProjectId(id); setActiveView('GANTT_DETAIL'); }} 
               onDeleteProject={handleDeleteProject}
               onUpdateProject={handleUpdateProject}
               onCreateRequested={() => setIsCreateProjectModalOpen(true)}
                onReschedule={(p) => setReschedulingProject(p)}
             />
           )}
           {activeView === 'SCHEDULE' && <OmDedySchedule user={user} users={users} setActiveView={setActiveView} />}
           {activeView === 'KANBAN' && (
             <KanbanView 
               projects={projects} 
               tasks={tasks}
               onOpenGantt={(id) => { setSelectedProjectId(id); setFocusedProjectId(id); setActiveView('GANTT_DETAIL'); }}
               onUpdateProject={handleUpdateProject}
             />
           )}
           {activeView === 'PERSONEL' && (
             <PersonelManagement isAdmin={isAdmin} 
               users={users} 
               projects={projects}
               currentUser={user}
               onRefresh={() => setRefreshKey(prev => prev + 1)}
             />
           )}
           {activeView === 'RESCHEDULE' && (
             <RescheduleRequestsView 
               requests={rescheduleRequests}
               onRefresh={() => {
                 fetchRescheduleRequests();
                 setRefreshKey(prev => prev + 1);
               }}
               user={user}
             />
           )}
           {activeView === 'AUDIT' && <AuditLogView logs={allAuditLogs} projects={projects} users={users} />}
           {activeView === 'GANTT_DETAIL' && (
             <GanttDetailView
               user={user}
               projectId={focusedProjectId}
               setFocusedProjectId={setFocusedProjectId}
               projects={projects}
               tasks={tasks}
               hierarchicalTasks={hierarchicalTasks}
               expandedRows={expandedRows}
               scale={scale}
               setRefreshKey={setRefreshKey}
               handleToggleExpand={handleToggleExpand}
               handleUpdateTask={handleUpdateTask}
               handleOpenAudit={handleOpenAudit}
               handleDeleteTask={requestDeleteTask}
               setScale={setScale}
               setTasks={setTasks}
             />
           )}
          </ErrorBoundary>
        </section>
      </main>

      <AnimatePresence>
        {isCreateProjectModalOpen && (
          <ErrorBoundary>
            <CreateProjectModal 
              user={user}
              users={users}
              onClose={() => setIsCreateProjectModalOpen(false)} 
              onSuccess={() => setRefreshKey(prev => prev + 1)} 
            />
          </ErrorBoundary>
        )}
        {showAuditLog && selectedTask && (
          <>
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowAuditLog(false)}
              className="fixed inset-0 bg-black/60 backdrop-blur-sm z-40"
            />
            <motion.div 
              initial={{ x: '100%' }}
              animate={{ x: 0 }}
              exit={{ x: '100%' }}
              className="fixed top-0 right-0 bottom-0 w-[450px] bg-slate-900 border-l border-slate-800 shadow-2xl z-50 overflow-y-auto"
            >
              <div className="p-6">
                <div className="flex items-center justify-between mb-8">
                  <div>
                    <h2 className="text-xl font-bold text-white flex items-center gap-2">
                      <History className="w-5 h-5 text-indigo-500" />
                      Immutable Audit Trail
                    </h2>
                    <p className="text-xs text-slate-500 mt-1">Governance Layer for {selectedTask.title}</p>
                  </div>
                  <button onClick={() => setShowAuditLog(false)} className="text-slate-400 hover:text-white">
                    <Plus className="w-6 h-6 rotate-45" />
                  </button>
                </div>

                <div className="space-y-6">
                    {auditLogs.map((log, i) => {
                      if (!log) return null;
                      const logId = log.id || `audit-${i}-${crypto.randomUUID()}`;
                      return (
                        <div key={logId} className="relative pl-8 pb-6 group">
                        {i !== (auditLogs || []).length - 1 && (
                          <div className="absolute left-[11px] top-6 bottom-0 w-[2px] bg-slate-800 group-last:hidden" />
                        )}
                        <div className="absolute left-0 top-1 w-6 h-6 rounded-full bg-slate-800 border-2 border-slate-700 flex items-center justify-center z-10">
                          <div className="w-2 h-2 rounded-full bg-indigo-500" />
                        </div>
                        
                        <div className="bg-slate-800/40 border border-slate-800 p-4 rounded-xl">
                          <div className="flex items-center justify-between mb-2">
                            <span className="text-xs font-bold text-indigo-400">{log.actor || 'System'}</span>
                            <span className="text-[10px] text-slate-500 font-mono">{log.created_at ? format(new Date(log.created_at), 'MM/dd HH:mm:ss') : 'N/A'}</span>
                          </div>
                          <p className="text-sm font-medium text-slate-200 mb-2">{log.action || 'Unknown Action'}</p>
                          
                          {log.old_payload && log.new_payload && (
                            <div className="mt-4 space-y-2 border-t border-slate-700/50 pt-4">
                              <p className="text-[10px] text-slate-500 font-black uppercase tracking-widest mb-2">Payload Comparison</p>
                              {(() => {
                                const oldP = log.old_payload as any;
                                const newP = log.new_payload as any;
                                const allKeys = Array.from(new Set([...Object.keys(oldP), ...Object.keys(newP)]));
                                
                                return allKeys.map((key, ki) => {
                                  if (key === 'updated_at' || key === 'id' || key === 'created_at') return null;
                                  if (JSON.stringify(oldP[key]) === JSON.stringify(newP[key])) return null;
                                  
                                  return (
                                    <div key={`payload-diff-${key}-${ki}`} className="flex flex-col gap-1 pb-2 border-b border-white/5 last:border-0">
                                      <span className="text-[10px] font-mono text-indigo-400/70">{key}</span>
                                      <div className="grid grid-cols-[1fr,auto,1fr] items-center gap-2">
                                        <div className="bg-rose-500/10 text-rose-400 text-[10px] p-1 rounded border border-rose-500/20 line-through opacity-60 truncate">
                                          {String(oldP[key] ?? 'null')}
                                        </div>
                                        <ArrowRight className="w-3 h-3 text-slate-600" />
                                        <div className="bg-emerald-500/10 text-emerald-400 text-[10px] p-1 rounded border border-emerald-500/20 font-bold italic truncate">
                                          {String(newP[key] ?? 'null')}
                                        </div>
                                      </div>
                                    </div>
                                  );
                                });
                              })()}
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </motion.div>
          </>
        )}
        {reschedulingProject && (
          <ProjectRescheduleModal 
            project={reschedulingProject}
            user={user}
            onClose={() => setReschedulingProject(null)}
            onSuccess={fetchData}
          />
        )}
        <ConfirmModal
          isOpen={!!itemToDelete}
          onClose={() => setItemToDelete(null)}
          onConfirm={confirmDelete}
          title="Konfirmasi Hapus"
          description="Apakah Anda yakin ingin menghapus data ini? Semua rincian di dalamnya juga akan ikut terhapus secara permanen. Aksi ini tidak dapat dibatalkan."
        />
        {taskModalData && (
          <CreateTaskModal 
            projectId={focusedProjectId!}
            parentId={taskModalData.parentId}
            onClose={() => setTaskModalData(null)}
            onSuccess={() => setRefreshKey(prev => prev + 1)}
            user={user}
          />
        )}
      </AnimatePresence>
    </div>
  );
}

// --- Modals ---

function ProjectRescheduleModal({ project, user, onClose, onSuccess }: { project: Project, user: any, onClose: () => void, onSuccess: () => void }) {
  const [newStart, setNewStart] = useState(project.start_date || format(new Date(), 'yyyy-MM-dd'));
  const [newEnd, setNewEnd] = useState(project.end_date || format(addDays(new Date(), 30), 'yyyy-MM-dd'));
  const [reason, setReason] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!reason.trim()) return;
    setLoading(true);
    try {
      // 1. Log the change
      await taskService.createProjectRescheduleLog({
        project_id: project.id,
        changed_by: user.name || user.email || 'Admin',
        old_start_date: project.start_date || 'N/A',
        old_end_date: project.end_date || 'N/A',
        new_start_date: newStart,
        new_end_date: newEnd,
        reason: reason
      });

      // 2. Update project
      await taskService.updateProject(project.id, {
        start_date: newStart,
        end_date: newEnd
      }, user.email || 'Admin');

      onSuccess();
      onClose();
    } catch (err: any) {
      console.error("Failed to reschedule project:", err);
      alert("Error: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
        className="absolute inset-0 bg-slate-950/80 backdrop-blur-md"
      />
      <motion.div 
        initial={{ opacity: 0, scale: 0.9, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 20 }}
        className="relative w-full max-w-lg w-[95%] mx-auto bg-slate-900 border border-white/10 rounded-3xl shadow-2xl overflow-hidden max-h-[90vh] overflow-y-auto"
      >
        <div className="p-8 border-b border-white/5 bg-gradient-to-br from-slate-900 to-slate-950">
          <div className="flex justify-between items-start mb-2">
            <div>
              <h3 className="text-2xl font-black text-white italic uppercase tracking-tighter">Reschedule Project</h3>
              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-[0.3em] mt-1">Timeline Governance Layer</p>
            </div>
            <button onClick={onClose} className="text-slate-500 hover:text-white transition-colors">
              <X className="w-6 h-6" />
            </button>
          </div>
          <div className="p-4 bg-indigo-500/5 rounded-2xl border border-indigo-500/10 mt-6">
            <h4 className="text-[10px] font-black text-indigo-400 uppercase tracking-widest mb-1">Target Infrastructure</h4>
            <p className="text-sm font-bold text-slate-200">{project.name}</p>
          </div>
        </div>

        <form onSubmit={handleSubmit} className="p-8 space-y-6">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
            <div className="space-y-4 opacity-50 grayscale pointer-events-none hidden sm:block">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest block">Current Timeline</label>
              <div className="space-y-3">
                <div className="flex flex-col gap-1.5 p-3 bg-slate-950/50 rounded-xl border border-white/5">
                  <span className="text-[8px] font-bold text-slate-600 uppercase">Start Date</span>
                  <span className="text-xs text-slate-400 font-mono italic">{project.start_date || 'N/A'}</span>
                </div>
                <div className="flex flex-col gap-1.5 p-3 bg-slate-950/50 rounded-xl border border-white/5">
                  <span className="text-[8px] font-bold text-slate-600 uppercase">End Date</span>
                  <span className="text-xs text-slate-400 font-mono italic">{project.end_date || 'N/A'}</span>
                </div>
              </div>
            </div>

            <div className="space-y-4">
              <label className="text-[10px] font-black text-indigo-500 uppercase tracking-widest block">New Timeline Configuration</label>
              <div className="space-y-4">
                <div className="flex flex-col gap-1.5">
                   <label className="text-[8px] font-bold text-slate-500 uppercase ml-1">Start Date</label>
                   <div className="relative group">
                     <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-600 group-focus-within:text-indigo-500 transition-colors" />
                     <input 
                       type="date"
                       value={newStart}
                       onChange={(e) => setNewStart(e.target.value)}
                       required
                       className="w-full bg-slate-900 border border-slate-700 text-white rounded-xl pl-10 pr-4 py-3 text-sm focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500/50 outline-none transition-all"
                     />
                   </div>
                </div>
                <div className="flex flex-col gap-1.5">
                   <label className="text-[8px] font-bold text-slate-500 uppercase ml-1">End Date</label>
                   <div className="relative group">
                     <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-600 group-focus-within:text-indigo-500 transition-colors" />
                     <input 
                       type="date"
                       value={newEnd}
                       onChange={(e) => setNewEnd(e.target.value)}
                       required
                       className="w-full bg-slate-900 border border-slate-700 text-white rounded-xl pl-10 pr-4 py-3 text-sm focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500/50 outline-none transition-all"
                     />
                   </div>
                </div>
              </div>
            </div>
          </div>

          <div className="space-y-3 pt-4 border-t border-white/5">
            <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-2">Justification / Reason</label>
            <textarea 
              required
              value={reason}
              onChange={(e) => setReason(e.target.value)}
              className="w-full bg-slate-950/50 border border-slate-800 rounded-2xl p-4 text-sm text-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500/50 outline-none transition-all placeholder:text-slate-700 min-h-[100px]"
              placeholder="Provide context for this schedule shift..."
            />
          </div>

          <div className="pt-4 flex gap-4">
            <button 
              type="button"
              onClick={onClose}
              className="flex-1 py-4 bg-slate-800/50 hover:bg-slate-800 text-slate-400 rounded-2xl text-[10px] font-black uppercase tracking-[0.2em] transition-all border border-white/5"
            >
              Abrupty Cancel
            </button>
            <button 
              type="submit"
              disabled={loading}
              className="flex-[2] py-4 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white text-[10px] font-black uppercase tracking-[0.2em] rounded-2xl shadow-xl shadow-indigo-600/20 transition-all active:scale-[0.98] disabled:opacity-50"
            >
              {loading ? 'Committing Changes...' : 'Execute Timeline Shift'}
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
}

function CreateTaskModal({ parentId, projectId, onClose, onSuccess, user, initialData }: { parentId: string | null, projectId: string, onClose: () => void, onSuccess: () => void, user: any, initialData?: Task }) {
  const [title, setTitle] = useState(initialData?.title || '');
  const [assignee, setAssignee] = useState(initialData?.assignee || '');
  const [fromDate, setFromDate] = useState(initialData?.start_time ? format(new Date(initialData.start_time), 'yyyy-MM-dd') : format(new Date(), 'yyyy-MM-dd'));
  const [toDate, setToDate] = useState(initialData?.end_time ? format(new Date(initialData.end_time), 'yyyy-MM-dd') : format(addDays(new Date(), 7), 'yyyy-MM-dd'));
  const [manHours, setManHours] = useState(initialData?.duration_hours || 0);
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (new Date(toDate) < new Date(fromDate)) {
      alert('To Date cannot be earlier than From Date');
      return;
    }

    setLoading(true);
    try {
      const taskData = {
        title,
        assignee: assignee || user?.name || user?.email || 'User',
        project_id: projectId,
        parent_id: parentId,
        start_time: new Date(fromDate).toISOString(),
        end_time: new Date(toDate).toISOString(),
        duration_hours: Number(manHours) || 0
      };

      if (initialData?.id) {
        await taskService.updateTask(initialData.id, taskData, user?.email || 'User');
      } else {
        await taskService.createTask(taskData, user?.email || 'User');
      }
      onSuccess();
      onClose();
    } catch (err: any) {
      alert('Failed to save task: ' + err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 z-[300] flex items-center justify-center p-4">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
        className="absolute inset-0 bg-slate-950/90 backdrop-blur-sm"
      />
      <motion.div 
        initial={{ opacity: 0, scale: 0.95, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.95, y: 20 }}
        className="relative w-full max-w-lg bg-slate-900 border border-white/10 rounded-3xl shadow-2xl overflow-hidden max-h-[90vh] overflow-y-auto"
      >
        <div className="p-6 border-b border-white/5 bg-gradient-to-br from-slate-900 to-slate-950">
          <h3 className="text-xl font-black text-white italic uppercase tracking-tight">
            {parentId ? 'Add Breakdown (L2)' : 'Add Phase (L1)'}
          </h3>
          <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mt-1">Infrastructure Partitioning</p>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-5">
          <div className="space-y-2">
            <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Title / Node Name</label>
            <input 
              required
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-2xl px-4 py-3 text-sm text-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500/50 outline-none transition-all"
              placeholder="Enter node title..."
            />
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">From Date</label>
              <input 
                type="date"
                required
                value={fromDate}
                onChange={(e) => setFromDate(e.target.value)}
                className="w-full bg-slate-900 border border-slate-700 rounded-2xl px-4 py-3 text-sm text-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500/50 outline-none transition-all"
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">To Date</label>
              <input 
                type="date"
                required
                value={toDate}
                onChange={(e) => setToDate(e.target.value)}
                className="w-full bg-slate-900 border border-slate-700 rounded-2xl px-4 py-3 text-sm text-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500/50 outline-none transition-all"
              />
            </div>
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Assignee / PIC</label>
              <input 
                value={assignee}
                onChange={(e) => setAssignee(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-2xl px-4 py-3 text-sm text-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500/50 outline-none transition-all"
                placeholder="Assign to..."
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Est. Man-Hours</label>
              <input 
                type="number"
                required
                value={manHours}
                onChange={(e) => setManHours(Number(e.target.value))}
                className="w-full bg-slate-950 border border-slate-800 rounded-2xl px-4 py-3 text-sm text-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500/50 outline-none transition-all"
              />
            </div>
          </div>

          <div className="pt-4 flex gap-4">
            <button 
              type="button"
              onClick={onClose}
              className="flex-1 py-3 bg-slate-800/50 hover:bg-slate-800 text-slate-400 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all border border-white/5"
            >
              Cancel
            </button>
            <button 
              type="submit"
              disabled={loading}
              className="flex-[2] py-3 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white text-[10px] font-black uppercase tracking-widest rounded-xl shadow-xl shadow-emerald-600/20 transition-all active:scale-[0.98] disabled:opacity-50"
            >
              {loading ? 'Creating...' : 'Finalize Persistence'}
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
}

function CreateProjectModal({ onClose, onSuccess, user, users }: { onClose: () => void, onSuccess: () => void, user: any, users: AppUser[] }) {
  const [title, setTitle] = useState('');
  const [pic, setPic] = useState(user.name || user.email || '');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [phases, setPhases] = useState([{ 
    id: crypto.randomUUID(),
    title: 'Phase 1', 
    assignee: pic || user.name || user.email || '',
    start_date: '',
    end_date: '',
    duration_hours: 0,
    subtasks: [{ 
      id: crypto.randomUUID(), 
      title: 'Initial Breakdown', 
      assignee: pic || user.name || user.email || '',
      start_date: '',
      end_date: '',
      duration_hours: 0
    }]
  }]);
  const [loading, setLoading] = useState(false);

  // Auto-Sync Level 1 to Project Header
  useEffect(() => {
    const validFromDates = phases.map(p => p.start_date).filter(Boolean);
    const validToDates = phases.map(p => p.end_date).filter(Boolean);

    if (validFromDates.length > 0) {
       const minDate = validFromDates.reduce((a, b) => a < b ? a : b);
       if (!startDate || minDate < startDate) setStartDate(minDate);
    }
    if (validToDates.length > 0) {
       const maxDate = validToDates.reduce((a, b) => a > b ? a : b);
       if (!endDate || maxDate > endDate) setEndDate(maxDate);
    }
  }, [phases, startDate, endDate]);

  // Allocation Validation
  const allocationStats = useMemo(() => {
    return phases.map(phase => {
      const capacity = phase.duration_hours || 0;
      const used = phase.subtasks.reduce((sum, sub) => sum + (sub.duration_hours || 0), 0);
      return { 
        capacity, 
        used, 
        isOver: used > capacity, 
        isUnder: used < capacity,
        isPerfect: Math.abs(used - capacity) < 0.01 && capacity > 0,
        remaining: capacity - used
      };
    });
  }, [phases]);

  const totalManHours = useMemo(() => {
    return phases.reduce((acc, phase) => acc + (phase.duration_hours || 0), 0);
  }, [phases]);

  const isAnyOverAllocated = allocationStats.some(s => s.isOver);

  const addPhase = () => {
    setPhases([...phases, { 
      id: crypto.randomUUID(),
      title: `Phase ${phases.length + 1}`, 
      assignee: pic || user.name || user.email || '',
      start_date: '',
      end_date: '',
      duration_hours: 0,
      subtasks: [{ 
        id: crypto.randomUUID(), 
        title: 'New Breakdown', 
        assignee: pic || user.name || user.email || '',
        start_date: '',
        end_date: '',
        duration_hours: 0
      }]
    }]);
  };

  const addSubtask = (phaseIndex: number) => {
    const newPhases = [...phases];
    newPhases[phaseIndex].subtasks.push({ 
      id: crypto.randomUUID(), 
      title: 'Sub-Breakdown', 
      assignee: pic || user.name || user.email || '',
      start_date: '',
      end_date: '',
      duration_hours: 0
    });
    setPhases(newPhases);
  };

  const handleCreate = async () => {
    if (!title || isAnyOverAllocated) return;
    setLoading(true);
    try {
      const finalPic = pic || user.name || user.email || 'Administrator';
      const actorName = user?.name || 'Fachrul Wisnu Novianto';
      const actorEmail = user.email || 'Administrator';
      
      const prj = await taskService.createProject({ 
        name: title, 
        status: ProjectStatus.ACTIVE,
        leader_email: user.email || null,
        pic_name: finalPic
      }, actorEmail);
      
      for (const phase of phases) {
        const h = 8;
        const m = 0;
        const durationH = parseFloat(String(phase.duration_hours ?? 0)) || 0;

        const phaseStartStr = phase.start_date || startDate || format(new Date(), 'yyyy-MM-dd');
        const phaseEndStr = phase.end_date || endDate || format(addDays(new Date(phaseStartStr), 7), 'yyyy-MM-dd');

        const phaseStart = new Date(`${phaseStartStr}T08:00:00`);
        const phaseEnd = new Date(`${phaseEndStr}T17:00:00`);
        
        const l1 = await taskService.createTask({
          title: phase.title || 'Untitled Phase',
          project_id: prj.id,
          assignee: (phase as any).assignee || finalPic,
          start_time: phaseStart.toISOString(),
          end_time: phaseEnd.toISOString(),
          start_hour: h,
          start_minute: m,
          duration_hours: durationH,
          duration_minutes: 0,
          created_by_name: user?.name || actorName
        }, actorEmail);

        for (const sub of phase.subtasks) {
          const subH = 8;
          const subM = 0;
          const durH = parseFloat(String(sub.duration_hours ?? 0)) || 0;

          const subStartStr = sub.start_date || phase.start_date || startDate || format(new Date(), 'yyyy-MM-dd');
          const subEndStr = sub.end_date || sub.start_date || phase.end_date || endDate || format(addDays(new Date(subStartStr), 7), 'yyyy-MM-dd');

          const start = new Date(`${subStartStr}T08:00:00`);
          const end = new Date(`${subEndStr}T17:00:00`);
          
          await taskService.createTask({
            title: sub.title || 'Untitled Sub-task',
            project_id: prj.id,
            parent_id: l1.id,
            assignee: (sub as any).assignee || finalPic,
            start_time: start.toISOString(),
            end_time: end.toISOString(),
            start_hour: subH,
            start_minute: subM,
            duration_hours: durH,
            duration_minutes: 0,
            created_by_name: user?.name || actorName
          }, actorEmail);
        }
      }

      onSuccess();
      onClose();
    } catch (err) {
      console.error("Project Creation Wizard failed:", err);
      alert("Failed to create project infrastructure.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[100] flex items-center justify-center p-4">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        className="bg-slate-900 border border-slate-800 rounded-3xl w-full max-w-7xl max-h-[95vh] overflow-hidden shadow-[0_0_50px_rgba(79,70,229,0.2)] flex flex-col"
      >
        <div className="p-6 border-b border-slate-800 bg-indigo-600/5 flex justify-between items-center">
          <div>
            <h2 className="text-xl font-black text-white tracking-tighter uppercase italic">
              Project <span className="text-indigo-500">Wizard</span>
            </h2>
            <p className="text-[10px] text-slate-500 mt-1 font-bold uppercase tracking-widest">Multi-Level Batch Provisioning</p>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-slate-800 rounded-full transition-colors">
            <Plus className="w-5 h-5 text-slate-500 rotate-45" />
          </button>
        </div>

        <div className="p-8 flex-1 overflow-x-auto overflow-y-auto space-y-8 scrollbar-hide">
          <div className="grid grid-cols-4 gap-8 max-w-6xl mx-auto">
            <div className="space-y-4">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Project Meta</label>
              <LocalInput 
                autoFocus
                value={title}
                onChange={v => setTitle(v)}
                placeholder="Project Name..."
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-slate-200 outline-none focus:border-indigo-500 transition-colors font-bold"
              />
            </div>
            <div className="space-y-4">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Lead PIC</label>
              <LocalInput 
                value={pic}
                onChange={v => setPic(v)}
                placeholder="Lead PIC Name..."
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-slate-200 outline-none focus:border-indigo-500 transition-colors font-bold"
              />
            </div>
            <div className="space-y-4">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Start Date</label>
              <input 
                type="date"
                value={startDate}
                onChange={e => setStartDate(e.target.value)}
                className="w-full bg-slate-900 border border-slate-700 text-white rounded-xl px-4 py-3 outline-none focus:border-indigo-500 transition-colors font-bold"
              />
            </div>
            <div className="space-y-4">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">End Date</label>
              <input 
                type="date"
                value={endDate}
                onChange={e => setEndDate(e.target.value)}
                className="w-full bg-slate-900 border border-slate-700 text-white rounded-xl px-4 py-3 outline-none focus:border-indigo-500 transition-colors font-bold"
              />
            </div>
          </div>

          <div className="space-y-6 min-w-[1100px]">
            <div className="flex justify-between items-center px-1">
              <label className="text-[11px] font-black text-indigo-500 uppercase tracking-widest">Work Breakdown Structure (WBS)</label>
              <button 
                onClick={addPhase}
                className="text-[9px] font-black uppercase text-indigo-400 hover:text-indigo-300 transition-colors bg-indigo-500/10 px-4 py-2 rounded-lg border border-indigo-500/20"
              >
                + Add Phase (L1)
              </button>
            </div>

            {/* Header Labels */}
            <div className="grid grid-cols-[1fr_130px_130px_80px_150px_80px] gap-4 px-6 mb-[-16px]">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-2">Phase / Task Name</label>
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">From Date</label>
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">To Date</label>
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">Man-Hours</label>
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">PIC</label>
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">Actions</label>
            </div>

            {phases.map((phase, pIdx) => {
              const stats = allocationStats[pIdx];
              return (
                <div key={phase.id || `wizard-phase-${pIdx}-${crypto.randomUUID()}`} className="bg-slate-950/40 border border-slate-800/60 rounded-2xl p-6 space-y-4 relative overflow-hidden group">
                  <div className="absolute top-0 left-0 w-1.5 h-full bg-indigo-600/30 group-hover:bg-indigo-500 transition-colors" />
                  
                  {/* L1 Header & Inputs */}
                  <div className="grid grid-cols-[1fr_130px_130px_80px_150px_80px] items-center gap-4 bg-slate-900/50 p-4 rounded-xl border border-slate-800/50 shadow-sm transition-all hover:bg-slate-900/70">
                    <div className="flex flex-col">
                      <LocalInput 
                        value={phase.title}
                        onChange={(v) => {
                          const newPhases = [...phases];
                          newPhases[pIdx].title = v;
                          setPhases(newPhases);
                        }}
                        className="w-full bg-slate-950 border border-slate-800 rounded-lg px-4 py-2.5 text-sm text-white focus:border-indigo-500 outline-none transition-all font-bold"
                        placeholder="Phase Title"
                      />
                    </div>

                    <div className="flex flex-col gap-1">
                      <input 
                        type="date"
                        value={phase.start_date}
                        min={startDate}
                        max={endDate || undefined}
                        onChange={(e) => {
                          const newPhases = [...phases];
                          newPhases[pIdx].start_date = e.target.value;
                          setPhases(newPhases);
                        }}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-[10px] text-white focus:border-indigo-500 outline-none transition-all"
                      />
                    </div>

                    <div className="flex flex-col gap-1">
                      <input 
                        type="date"
                        value={phase.end_date}
                        min={phase.start_date || startDate}
                        max={endDate || undefined}
                        onChange={(e) => {
                          const newPhases = [...phases];
                          newPhases[pIdx].end_date = e.target.value;
                          setPhases(newPhases);
                        }}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-[10px] text-white focus:border-indigo-500 outline-none transition-all"
                      />
                    </div>

                    <div className="flex flex-col items-center gap-1">
                      <input 
                        type="number" min="0" step="0.5"
                        value={phase.duration_hours}
                        onChange={(e) => {
                          const newPhases = [...phases];
                          newPhases[pIdx].duration_hours = Math.max(0, parseFloat(e.target.value) || 0);
                          setPhases(newPhases);
                        }}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-xs text-indigo-400 text-center font-black outline-none focus:border-indigo-500"
                      />
                      <span className="text-[8px] font-black text-slate-600 uppercase tracking-tight">
                         ≈ {formatWorkday(phase.duration_hours)}
                      </span>
                    </div>

                    <div className="flex flex-col">
                      <LocalInput 
                        value={(phase as any).assignee}
                        onChange={(v) => {
                          const newPhases = [...phases];
                          (newPhases[pIdx] as any).assignee = v;
                          setPhases(newPhases);
                        }}
                        className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-[10px] text-slate-300 text-center outline-none focus:border-indigo-500"
                        placeholder="PIC Name"
                      />
                    </div>

                    <div className="flex justify-center gap-3">
                      <button 
                        onClick={() => addSubtask(pIdx)}
                        className="p-2 hover:bg-indigo-500/10 text-indigo-500/40 hover:text-indigo-400 rounded-lg transition-all"
                        title="Add Breakdown"
                      >
                        <Plus className="w-5 h-5" />
                      </button>
                      <button 
                        onClick={() => {
                          if (confirm('Apakah Anda yakin ingin menghapus data ini? Semua rincian di dalamnya juga akan terhapus secara permanen.')) {
                            const newPhases = [...phases];
                            newPhases.splice(pIdx, 1);
                            setPhases(newPhases);
                          }
                        }}
                        className="p-2 hover:bg-rose-500/10 text-slate-700 hover:text-rose-500 rounded-lg transition-all"
                        title="Delete Phase"
                      >
                        <Trash2 className="w-5 h-5" />
                      </button>
                    </div>
                  </div>

                  {/* Allocation Validation Message */}
                  <div className="px-4">
                    {stats.isOver && (
                      <p className="text-[10px] font-black text-rose-500 uppercase tracking-widest flex items-center gap-2 bg-rose-500/10 py-2 px-3 rounded-lg border border-rose-500/20">
                        <AlertTriangle className="w-3 h-3" />
                        ⚠️ Over-Allocated: Breakdown melebihi durasi Task Utama! (Kekurangan {formatWorkday(Math.abs(stats.remaining))})
                      </p>
                    )}
                    {stats.isUnder && (
                      <p className="text-[10px] font-black text-amber-500 uppercase tracking-widest flex items-center gap-2 bg-amber-500/10 py-2 px-3 rounded-lg border border-amber-500/20">
                        <AlertTriangle className="w-3 h-3" />
                        ⚠️ Under-Allocated: Sisa waktu {formatWorkday(stats.remaining)} belum dialokasikan
                      </p>
                    )}
                    {stats.isPerfect && (
                      <p className="text-[10px] font-black text-emerald-500 uppercase tracking-widest flex items-center gap-2 bg-emerald-500/5 py-2 px-3 rounded-lg border border-emerald-500/20">
                        <Plus className="w-3 h-3 text-emerald-500" />
                        ✅ Alokasi Waktu Pas
                      </p>
                    )}
                    {stats.capacity === 0 && (
                       <p className="text-[9px] font-bold text-slate-600 italic uppercase">Tentukan durasi untuk menghitung validasi alokasi.</p>
                    )}
                  </div>

                  <div className="space-y-3">
                    <div className="flex justify-between items-center px-4">
                      <p className="text-[9px] font-black text-slate-400 uppercase tracking-[0.2em] flex items-center gap-2">
                        <span className="w-1.5 h-1.5 rounded-full bg-indigo-500/50" />
                        Execution Breakdown
                      </p>
                    </div>

                    <div className="space-y-1.5">
                      {phase.subtasks.map((sub, sIdx) => (
                        <div key={sub.id} className="grid grid-cols-[1fr_130px_130px_80px_150px_80px] items-center gap-4 group/sub hover:bg-slate-900/30 p-2 rounded-lg transition-colors mx-2">
                          {/* Col 1 with Indentation Built-in */}
                          <div className="flex items-center gap-3 pl-8">
                            <span className="text-indigo-500/20 font-black text-xl select-none group-hover/sub:text-indigo-500/40">↳</span>
                            <LocalInput 
                              value={sub.title}
                              placeholder="Activity description..."
                              onChange={(v) => {
                                const newPhases = [...phases];
                                newPhases[pIdx].subtasks[sIdx].title = v;
                                setPhases(newPhases);
                              }}
                              className="w-full bg-transparent border-b border-white/5 text-xs text-slate-400 py-1.5 outline-none focus:border-indigo-500/50 transition-all font-medium"
                            />
                          </div>

                          <div className="flex justify-center">
                            <input 
                              type="date"
                              value={sub.start_date}
                              min={phase.start_date}
                              max={phase.end_date || undefined}
                              onChange={(e) => {
                                const val = e.target.value;
                                if (phase.start_date && val < phase.start_date) {
                                  alert("Tanggal mulai child task tidak boleh mendahului tanggal pada task!");
                                  return;
                                }
                                const newPhases = [...phases];
                                newPhases[pIdx].subtasks[sIdx].start_date = val;
                                setPhases(newPhases);
                              }}
                              className="w-full bg-slate-900 border border-slate-700 rounded-lg px-1 py-1.5 text-[9px] text-white focus:border-indigo-500/50 outline-none transition-all"
                            />
                          </div>

                          <div className="flex justify-center">
                            <input 
                              type="date"
                              value={sub.end_date}
                              min={sub.start_date || phase.start_date}
                              max={phase.end_date || undefined}
                              onChange={(e) => {
                                const val = e.target.value;
                                if (phase.end_date && val > phase.end_date) {
                                  alert("Tanggal selesai child task tidak boleh melebihi task utama!");
                                  return;
                                }
                                const newPhases = [...phases];
                                newPhases[pIdx].subtasks[sIdx].end_date = val;
                                setPhases(newPhases);
                              }}
                              className="w-full bg-slate-900 border border-slate-700 rounded-lg px-1 py-1.5 text-[9px] text-white focus:border-indigo-500/50 outline-none transition-all"
                            />
                          </div>

                          <div className="flex justify-center">
                            <input 
                              type="number" min="0" step="0.5"
                              value={sub.duration_hours}
                              onChange={(e) => {
                                const newPhases = [...phases];
                                newPhases[pIdx].subtasks[sIdx].duration_hours = Math.max(0, parseFloat(e.target.value) || 0);
                                setPhases(newPhases);
                              }}
                              className="w-16 bg-slate-900 border border-slate-700 rounded-md px-1 py-1 text-[10px] text-indigo-400/80 text-center outline-none focus:border-indigo-500/50 font-black"
                            />
                          </div>

                          <div className="flex justify-center">
                            <LocalInput 
                              value={(sub as any).assignee}
                              onChange={(v) => {
                                const newPhases = [...phases];
                                (newPhases[pIdx].subtasks[sIdx] as any).assignee = v;
                                setPhases(newPhases);
                              }}
                              placeholder="Assignee"
                              className="w-full bg-slate-950 border border-slate-800 rounded px-2 py-1 text-[9px] text-slate-500 text-center outline-none focus:border-indigo-500/50"
                            />
                          </div>

                          <div className="flex justify-center">
                            <button 
                              onClick={() => {
                                if (confirm('Apakah Anda yakin ingin menghapus data ini? Semua rincian di dalamnya juga akan terhapus secara permanen.')) {
                                  const newPhases = [...phases];
                                  newPhases[pIdx].subtasks.splice(sIdx, 1);
                                  setPhases(newPhases);
                                }
                              }}
                              className="p-1.5 opacity-0 group-hover/sub:opacity-100 hover:bg-rose-500/10 text-slate-700 hover:text-rose-500 rounded transition-all"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        <div className="p-8 border-t border-slate-800 bg-slate-950/50 flex justify-between items-center px-10">
          <p className="text-[10px] text-slate-500 font-bold uppercase tracking-[0.2em] flex items-center gap-6">
            <span className="flex items-center gap-2">
              <span className="w-2.5 h-2.5 rounded-full bg-indigo-500 animate-pulse outline outline-offset-2 outline-indigo-500/20" />
              {phases.length} Phases
            </span>
            <span className="flex items-center gap-2">
               <span className="w-2.5 h-2.5 rounded-full bg-slate-700" />
               <span className="text-indigo-400 font-black">Total Man-Hours: {totalManHours.toFixed(1)} Jam</span>
               <span className="text-slate-600 font-medium ml-1">≈ {formatWorkday(totalManHours)}</span>
            </span>
            {isAnyOverAllocated && (
               <span className="text-rose-500 font-black animate-pulse">⚠️ Resolusi Over-Alokasi diperlukan</span>
            )}
          </p>
          <div className="flex gap-4">
            <button 
              disabled={loading}
              onClick={onClose}
              className="px-8 py-3 bg-slate-800 hover:bg-slate-700 text-slate-500 font-black uppercase text-[10px] tracking-widest rounded-xl transition-all"
            >
              Cancel
            </button>
            <button 
              disabled={loading || !title || isAnyOverAllocated}
              onClick={handleCreate}
              className="bg-indigo-600 hover:bg-indigo-500 disabled:opacity-30 disabled:grayscale text-white px-10 py-4 rounded-xl text-xs font-black uppercase tracking-widest transition-all shadow-[0_10px_30px_rgba(79,70,229,0.3)] hover:shadow-[0_15px_40px_rgba(79,70,229,0.4)] flex items-center gap-3 disabled:cursor-not-allowed"
            >
              {loading ? (
                <div className="w-5 h-5 border-2 border-white/20 border-t-white rounded-full animate-spin" />
              ) : (
                <>
                  <Rocket className="w-5 h-5" />
                  Spawn Infrastructure
                </>
              )}
            </button>
          </div>
        </div>
      </motion.div>
    </div>
  );
}

function BatchManualEntryModal({ onClose, onSuccess, users }: { onClose: () => void, onSuccess: () => void, users: AppUser[] }) {
  const [rows, setRows] = useState([{ pic_name: '', status: 'WFO', schedule_date: format(new Date(), 'yyyy-MM-dd') }]);
  const [pasteData, setPasteData] = useState('');
  const [loading, setLoading] = useState(false);

  const addRow = () => {
    const lastRow = rows[rows.length - 1];
    setRows([...rows, { 
      pic_name: lastRow?.pic_name || '', 
      status: 'WFO', 
      schedule_date: lastRow?.schedule_date || format(new Date(), 'yyyy-MM-dd') 
    }]);
  };

  const handlePaste = () => {
    if (!pasteData.trim()) return;
    const lines = pasteData.trim().split('\n');
    const newRows = lines.map(line => {
      const parts = line.split('\t');
      if (parts.length >= 2) {
        const pic_name = parts[0].trim();
        const status = (parts[1]?.trim().toUpperCase() || 'WFO');
        let schedule_date = parts[2]?.trim() || format(new Date(), 'yyyy-MM-dd');
        
        // Basic date normalization if it looks like DD/MM/YYYY
        if (schedule_date.includes('/')) {
          const dateParts = schedule_date.split('/');
          if (dateParts.length === 3) {
            const [d, m, y] = dateParts;
            schedule_date = `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
          }
        }
        
        return { pic_name, status, schedule_date };
      }
      return null;
    }).filter(Boolean) as any[];

    if (newRows.length > 0) {
      setRows([...rows, ...newRows]);
      setPasteData('');
    }
  };

  const handleSave = async () => {
    setLoading(true);
    try {
      const payload = rows.filter(r => r.pic_name && r.schedule_date);
      if (payload.length === 0) return;
      
      // Chunk processing for safety
      for (let i = 0; i < payload.length; i += 50) {
        await taskService.upsertSchedules(payload.slice(i, i + 50));
      }
      
      onSuccess();
      onClose();
    } catch (err) {
      console.error('Batch upsert failed:', err);
      alert('Gagal menyimpan data massal.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 z-[110] flex items-center justify-center p-4">
      <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} onClick={onClose} className="absolute inset-0 bg-slate-950/90 backdrop-blur-md" />
      <motion.div 
        initial={{ opacity: 0, scale: 0.95, y: 30 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        className="relative w-full max-w-4xl bg-slate-900 border border-slate-800 rounded-[2.5rem] shadow-2xl overflow-hidden max-h-[90vh] flex flex-col"
      >
        <div className="p-8 border-b border-slate-800 flex items-center justify-between bg-slate-900/50 backdrop-blur-md">
          <div>
            <h2 className="text-3xl font-black text-white italic tracking-tighter uppercase leading-none">Batch Manual Entry</h2>
            <p className="text-[10px] font-black text-slate-500 uppercase tracking-[0.4em] mt-3 flex items-center gap-2">
              <Copy className="w-3 h-3 text-indigo-500" /> Massive Schedule Update
            </p>
          </div>
          <button onClick={onClose} className="p-3 hover:bg-slate-800 rounded-2xl text-slate-400 transition-all">
            <X className="w-6 h-6" />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-8 space-y-8 custom-scrollbar">
          {/* Paste Area */}
          <div className="space-y-3">
            <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Paste from Excel (Tab-Separated: Name \t Status \t Date)</label>
            <div className="flex gap-4">
              <textarea 
                value={pasteData}
                onChange={(e) => setPasteData(e.target.value)}
                placeholder="Paste multi-line data here..."
                className="flex-1 bg-slate-950/50 border border-slate-800 rounded-2xl p-4 text-xs font-mono text-slate-300 min-h-[100px] outline-none focus:border-indigo-500/50 transition-all"
              />
              <button 
                onClick={handlePaste}
                className="px-6 bg-slate-800 hover:bg-slate-700 text-white rounded-2xl text-[10px] font-black uppercase tracking-widest transition-all"
              >
                Parse Paste
              </button>
            </div>
          </div>

          {/* Manual Table */}
          <div className="space-y-4">
            <div className="flex items-center justify-between px-1">
              <h3 className="text-sm font-black text-white uppercase tracking-widest">Entry Items</h3>
              <button onClick={addRow} className="flex items-center gap-2 px-3 py-1.5 bg-indigo-500/10 hover:bg-indigo-500/20 text-indigo-400 text-[10px] font-black uppercase tracking-widest rounded-lg border border-indigo-500/20 transition-all">
                <Plus className="w-3 h-3" /> Tambah Baris
              </button>
            </div>

            <div className="space-y-2 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
              {rows.map((row, idx) => (
                <div key={idx} className="grid grid-cols-[1fr_150px_180px_50px] gap-3 items-center bg-slate-950/30 p-2 rounded-xl border border-white/[0.02]">
                  <div className="relative group">
                    <select 
                      value={row.pic_name}
                      onChange={(e) => {
                        const newRows = [...rows];
                        newRows[idx].pic_name = e.target.value;
                        setRows(newRows);
                      }}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-xs text-white focus:border-indigo-500/50 outline-none"
                    >
                      <option value="">Select PIC</option>
                      {users.map(u => <option key={u.id} value={u.name}>{u.name}</option>)}
                    </select>
                  </div>
                  <select 
                    value={row.status}
                    onChange={(e) => {
                      const newRows = [...rows];
                      newRows[idx].status = e.target.value;
                      setRows(newRows);
                    }}
                    className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-xs text-white focus:border-indigo-500/50 outline-none font-bold"
                  >
                    <option value="WFO">WFO</option>
                    <option value="WFH">WFH</option>
                    <option value="WFC">WFC</option>
                    <option value="LIBUR">LIBUR</option>
                  </select>
                  <input 
                    type="date"
                    value={row.schedule_date}
                    onChange={(e) => {
                      const newRows = [...rows];
                      newRows[idx].schedule_date = e.target.value;
                      setRows(newRows);
                    }}
                    className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-xs text-white focus:border-indigo-500/50 outline-none font-bold"
                  />
                  <button 
                    onClick={() => setRows(rows.filter((_, i) => i !== idx))}
                    className="text-slate-600 hover:text-rose-500 p-2"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>
              ))}
            </div>
          </div>
        </div>

        <div className="p-8 border-t border-slate-800 bg-slate-900 flex items-center justify-end gap-4">
          <button onClick={onClose} className="px-6 py-3 bg-slate-800 text-slate-400 text-[10px] font-black uppercase tracking-widest rounded-xl transition-all">Cancel</button>
          <button 
            onClick={handleSave}
            disabled={loading || rows.length === 0}
            className="px-10 py-3 bg-emerald-600 hover:bg-emerald-500 text-white text-[10px] font-black uppercase tracking-widest rounded-xl shadow-lg shadow-emerald-500/20 transition-all active:scale-95"
          >
            {loading ? 'Processing...' : 'Simpan Semua'}
          </button>
        </div>
      </motion.div>
    </div>
  );
}

// --- View Components ---

function DashboardStats({ tasks, projects }: { tasks: Task[], projects: Project[] }) {
  const stats = useMemo(() => {
    const totalProjects = projects.length;
    const totalTasks = tasks.filter(t => t.parent_id === null).length;
    const totalChildTasks = tasks.filter(t => t.parent_id !== null).length;
    const totalManHours = tasks
      .filter(t => t.parent_id === null)
      .reduce((sum, t) => sum + (Number(t.duration_hours) || 0), 0);
    
    return { totalProjects, totalTasks, totalChildTasks, totalManHours };
  }, [tasks, projects]);

  return (
    <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
      {[
        { label: 'Total Project', value: stats.totalProjects, icon: FolderKanban, color: 'emerald' },
        { label: 'Total Task', value: stats.totalTasks, icon: Layers, color: 'sky' },
        { label: 'Total Child Task', value: stats.totalChildTasks, icon: Activity, color: 'indigo' },
        { label: 'Total Man Hour', value: stats.totalManHours.toFixed(1), icon: Clock, color: 'amber' },
      ].map((stat, i) => (
        <div key={`dashboard-stat-${stat.label}-${i}`} className="bg-slate-900/60 backdrop-blur-xl border border-white/5 p-5 rounded-2xl relative overflow-hidden group hover:border-indigo-500/30 transition-all shadow-2xl">
          <div className="flex items-center justify-between relative z-10">
            <div>
              <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest mb-1">{stat.label}</p>
              <h3 className="text-2xl font-black text-white italic tracking-tighter">{stat.value}</h3>
            </div>
            <div className={cn(
              "w-10 h-10 rounded-xl flex items-center justify-center transition-transform group-hover:scale-110",
              stat.color === 'emerald' ? "bg-emerald-500/10 text-emerald-400" :
              stat.color === 'sky' ? "bg-sky-500/10 text-sky-400" :
              stat.color === 'indigo' ? "bg-indigo-500/10 text-indigo-400" :
              "bg-amber-500/10 text-amber-400"
            )}>
              <stat.icon className="w-5 h-5" />
            </div>
          </div>
          <div className={cn(
            "absolute bottom-0 left-0 h-1 transition-all",
            stat.color === 'emerald' ? "bg-emerald-500/20 group-hover:bg-emerald-500" :
            stat.color === 'sky' ? "bg-sky-500/20 group-hover:bg-sky-500" :
            stat.color === 'indigo' ? "bg-indigo-500/20 group-hover:bg-indigo-500" :
            "bg-amber-500/20 group-hover:bg-amber-500"
          )} style={{ width: '100%' }} />
        </div>
      ))}
    </div>
  );
}

function PortfolioDashboard({ projects, tasks, loading, onOpenProject, onDeleteProject, onUpdateProject, onCreateRequested, onReschedule }: { 
  projects: Project[], 
  tasks: Task[],
  loading: boolean,
  onOpenProject: (id: string) => void,
  onDeleteProject: (id: string) => void,
  onUpdateProject: (id: string, updates: Partial<Project>) => void,
  onCreateRequested: () => void,
  onReschedule: (p: Project) => void
}) {
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [notif, setNotif] = useState<string | null>(null);
  const safeProjects = projects || [];

  const handleExportProjects = () => {
    const exportData = safeProjects.map(p => ({
      'Project Name': p.name || 'Untitled',
      'Lead PIC': p.pic_name || p.leader_email || 'Unassigned',
      'Start Date': p.start_date || 'N/A',
      'End Date': p.end_date || 'N/A',
      'Status': p.status || 'Unknown'
    }));
    handleExcelExport(exportData, 'Projects_Summary');
  };

  return (
    <div className="space-y-6">
      <SuccessNotification show={!!notif} message={notif || ''} onClose={() => setNotif(null)} />
      
      <div className="flex items-center justify-between gap-4 mb-2">
        <div>
          <h2 className="text-2xl font-black text-white tracking-tighter italic uppercase">Om Dedy Portfolio</h2>
          <p className="text-xs text-slate-500 font-bold uppercase tracking-[0.2em] mt-1">Global Project Tracking System</p>
        </div>
        <div className="flex items-center gap-3">
          <button 
            onClick={handleExportProjects}
            className="bg-slate-800 hover:bg-slate-700 text-slate-300 px-4 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest flex items-center gap-2 border border-slate-700 transition-all active:scale-95"
          >
            <Download className="w-4 h-4" /> Export to Excel
          </button>
          <button 
            onClick={onCreateRequested}
            className="bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white px-4 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest shadow-lg shadow-indigo-500/20 flex items-center gap-2 transition-all active:scale-95"
          >
            <Plus className="w-4 h-4" /> New Project
          </button>
        </div>
      </div>
      
      {!loading && <DashboardStats tasks={tasks} projects={projects} />}
      
      {loading ? (
        <div className="h-64 flex flex-col items-center justify-center gap-4 bg-slate-900/20 rounded-3xl border border-slate-800/50">
           <div className="w-8 h-8 border-2 border-indigo-500 border-t-transparent rounded-full animate-spin" />
           <p className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-500">Retrieving Portfolio Data</p>
        </div>
      ) : (
        <>
          <ConfirmModal 
            isOpen={!!deleteId}
            onClose={() => setDeleteId(null)}
            onConfirm={() => {
              if (deleteId) {
                onDeleteProject(deleteId);
                setNotif('Project successfully decommissioned');
              }
            }}
            title="Archive Project?"
            description="Are you sure you want to remove this project? All associated tasks and audit trails will be archived."
          />

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {(safeProjects || []).map((p, i) => {
              if (!p) return null;
              const projectUniqueId = p.id || `project-fallback-${i}-${crypto.randomUUID()}`;
              return (
                <motion.div
                  key={`portfolio-project-${projectUniqueId}`}
                  whileHover={{ y: -4 }}
                  className="bg-slate-900 border border-slate-800 rounded-2xl p-6 cursor-pointer hover:border-indigo-500/50 transition-all group relative overflow-hidden"
                >
              <div className="absolute top-4 right-4 flex gap-2 opacity-0 group-hover:opacity-100 transition-opacity z-10">
                <button 
                  onClick={(e) => {
                    e.stopPropagation();
                    onReschedule(p);
                  }}
                  className="p-2 bg-slate-800 hover:bg-indigo-500/20 text-slate-400 hover:text-indigo-400 rounded-lg transition-colors border border-slate-700 hover:border-indigo-500/50"
                  title="Reschedule Project"
                >
                  <Calendar className="w-3.5 h-3.5" />
                </button>
                <button 
                  onClick={(e) => {
                    e.stopPropagation();
                    setDeleteId(p.id);
                  }}
                  className="p-2 bg-slate-800 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 rounded-lg transition-colors border border-slate-700 hover:border-rose-500/50"
                  title="Archive Project"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              </div>

                  <div onClick={() => onOpenProject(p.id)}>
                <div className="flex justify-between items-start mb-6">
                  <div className="w-12 h-12 bg-slate-800 rounded-xl flex items-center justify-center group-hover:bg-indigo-500/10 transition-colors">
                    <FolderKanban className="w-6 h-6 text-slate-400 group-hover:text-indigo-400" />
                  </div>
                  <div onClick={(e) => e.stopPropagation()}>
                    <ProjectStatusSelector status={p.status} onUpdate={(v) => onUpdateProject(p.id, { status: v })} />
                  </div>
                </div>
                <input 
                  defaultValue={p.name || ''}
                  onClick={(e) => e.stopPropagation()}
                  onBlur={(e) => {
                    if (e.target.value !== p.name) {
                      onUpdateProject(p.id, { name: e.target.value });
                      setNotif('Project title synchronized');
                    }
                  }}
                  className="text-lg font-bold text-white mb-2 bg-transparent border-none outline-none focus:text-indigo-400 w-full"
                  placeholder="Untitled Project"
                />
                
                <div className="flex flex-col gap-1.5 mb-2">
                  <div className="flex items-center gap-2 text-[9px] font-bold text-slate-500 uppercase tracking-widest">
                    <UserIcon className="w-3 h-3 text-indigo-500" />
                    <span>PIC: <span className="text-slate-300">{p.pic_name || p.leader_email || 'Unassigned'}</span></span>
                  </div>
                  <div className="flex items-center gap-2 text-[9px] font-bold text-slate-500 uppercase tracking-widest">
                    <Clock className="w-3 h-3 text-slate-600" />
                    <span>Timeline: <span className="text-slate-300 font-mono tracking-tighter">
                      {p.start_date ? format(new Date(p.start_date), 'MMM dd') : '??'} - {p.end_date ? format(new Date(p.end_date), 'MMM dd, yyyy') : '??'}
                    </span></span>
                  </div>
                </div>

                <div className="flex items-center justify-between text-[10px] font-black uppercase tracking-widest text-slate-600 mt-4 pt-4 border-t border-slate-800 group-hover:text-indigo-400 transition-colors">
                  <span>View Full Timeline</span>
                  <ExternalLink className="w-4 h-4 opacity-0 group-hover:opacity-100 transition-all transform translate-x-2 group-hover:translate-x-0 text-indigo-400" />
                </div>
              </div>
            </motion.div>
          );
        })}
        
        <button 
          onClick={onCreateRequested}
          className="border-2 border-dashed border-slate-800 rounded-2xl flex flex-col items-center justify-center p-6 text-slate-500 hover:text-indigo-400 hover:border-indigo-500/50 transition-all gap-3 bg-slate-900/10 group"
        >
          <div className="w-12 h-12 rounded-2xl bg-slate-800 group-hover:bg-gradient-to-br group-hover:from-indigo-600 group-hover:to-violet-600 flex items-center justify-center transition-all group-hover:shadow-[0_0_20px_rgba(79,70,229,0.3)]">
            <Plus className="w-8 h-8 group-hover:rotate-90 group-hover:text-white transition-all" />
          </div>
          <span className="text-xs font-black uppercase tracking-widest group-hover:text-indigo-400 transition-colors">Create New Project</span>
        </button>
      </div>
    </>
  )}
</div>
  );
}

const SCHEDULE_STATUS_COLORS: Record<string, string> = {
  'WFO': 'text-emerald-400 bg-emerald-500/10 border-emerald-500/20 shadow-[0_0_8px_rgba(16,185,129,0.1)]',
  'WFH': 'text-indigo-400 bg-indigo-500/10 border-indigo-500/20 shadow-[0_0_8px_rgba(99,102,241,0.1)]',
  'WFC': 'text-amber-400 bg-amber-500/10 border-amber-500/20 shadow-[0_0_8px_rgba(245,158,11,0.1)]',
  'A2': 'text-amber-400 bg-amber-500/10 border-amber-500/20 shadow-[0_0_8px_rgba(245,158,11,0.1)]',
  'LIBUR': 'text-rose-400 bg-rose-500/10 border-rose-500/20 shadow-[0_0_8px_rgba(244,63,94,0.1)]',
  'LATE': 'text-rose-400 bg-rose-500/10 border-rose-500/20 shadow-[0_0_8px_rgba(244,63,94,0.1)]',
};

function OmDedySchedule({ user, users, setActiveView }: { user: any, users: AppUser[], setActiveView: (view: AppView) => void }) {
  const [currentMonth, setCurrentMonth] = useState(new Date());
  const [schedules, setSchedules] = useState<Schedule[]>([]);
  const [loading, setLoading] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [showBatchModal, setShowBatchModal] = useState(false);
  const [editingCell, setEditingCell] = useState<{ pic: string, date: string } | null>(null);
  const [rescheduleRequests, setRescheduleRequests] = useState<any[]>([]);
  const [showRescheduleRequests, setShowRescheduleRequests] = useState(false);
  const [requestModal, setRequestModal] = useState<{ 
    pic: string, 
    date: string, 
    currentStatus: string, 
    newStatus: string,
    isSwap?: boolean,
    swapDate?: string,
    swapStatus?: string,
    swapCurrentStatus?: string
  } | null>(null);

  const currentUserProfile = useMemo(() => {
    return users.find(u => u.email?.toLowerCase() === user?.email?.toLowerCase());
  }, [users, user]);

  const userAccess = (currentUserProfile?.access_level || '').toLowerCase().trim();
  const isAdmin = userAccess === 'superadmin' || userAccess === 'admin';
  const fullName = currentUserProfile?.name || '';

  const monthStart = startOfMonth(currentMonth);
  const monthEnd = endOfMonth(currentMonth);
  const days = eachDayOfInterval({ start: monthStart, end: monthEnd });

  const stats = useMemo(() => {
    const counts = { WFO: 0, WFH: 0, CUTI_LIBUR: 0, WFC: 0 };
    schedules.forEach(s => {
      // Only count if it's in the current month's days
      const dateKey = s.schedule_date; // already YYYY-MM-DD
      const date = new Date(dateKey);
      if (date >= monthStart && date <= monthEnd) {
        const st = s.status.toUpperCase();
        if (st === 'WFO') counts.WFO++;
        else if (st === 'WFH') counts.WFH++;
        else if (st === 'WFC' || st === 'A2') counts.WFC++;
        else if (st === 'CUTI' || st === 'LIBUR') counts.CUTI_LIBUR++;
      }
    });
    return counts;
  }, [schedules, monthStart, monthEnd]);

  const picStats = useMemo(() => {
    const statsMap: Record<string, { WFO: number, WFH: number, TOTAL: number }> = {};
    const list = Array.from(new Set(schedules.map(s => s.pic_name))).sort();
    
    list.forEach(p => statsMap[p] = { WFO: 0, WFH: 0, TOTAL: 0 });
    
    schedules.forEach(s => {
      const d = new Date(s.schedule_date);
      if (d >= monthStart && d <= monthEnd && statsMap[s.pic_name]) {
        const st = (s.status || '').toUpperCase();
        if (st === 'WFO') statsMap[s.pic_name].WFO++;
        else if (st === 'WFH') statsMap[s.pic_name].WFH++;
        statsMap[s.pic_name].TOTAL = statsMap[s.pic_name].WFO + statsMap[s.pic_name].WFH;
      }
    });
    return statsMap;
  }, [schedules, monthStart, monthEnd]);

  const fetchSchedules = async () => {
    try {
      setLoading(true);
      const data = await taskService.getSchedules(currentMonth);
      setSchedules(data);
      if (user.access_level !== 'PIC') {
        const reqs = await taskService.getRescheduleRequests();
        setRescheduleRequests(reqs);
      }
    } catch (err) {
      console.error('Failed to fetch schedules:', err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchSchedules();
    
    // Subscribe to schedule changes
    const scheduleChannel = supabase.channel('schedules-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'schedules' }, () => {
        fetchSchedules();
      })
      .subscribe();

    const requestChannel = supabase.channel('requests-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'reschedule_requests' }, () => {
        fetchSchedules();
      })
      .subscribe();
      
    return () => { 
      supabase.removeChannel(scheduleChannel); 
      supabase.removeChannel(requestChannel);
    };
  }, [currentMonth]);

  const pics = useMemo(() => {
    // Collect all unique PICs from schedules
    const list = Array.from(new Set(schedules.map(s => s.pic_name))).sort();
    return list;
  }, [schedules]);

  const scheduleGrid = useMemo(() => {
    const grid: Record<string, Record<string, string>> = {};
    schedules.forEach(s => {
      const dateKey = format(new Date(s.schedule_date), 'yyyy-MM-dd');
      if (!grid[s.pic_name]) grid[s.pic_name] = {};
      grid[s.pic_name][dateKey] = s.status.toUpperCase();
    });
    return grid;
  }, [schedules]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsImporting(true);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        
        const rawJsonAoA: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const jsonData: any[] = XLSX.utils.sheet_to_json(ws);

        let upsertPayload: Partial<Schedule>[] = [];

        // 1. Detect if it's a FLAT format (name, status, date)
        const firstRow = jsonData[0] || {};
        const keys = Object.keys(firstRow).map(k => k.toLowerCase());
        const flatMapping = {
          name: Object.keys(firstRow).find(k => ['name', 'pic'].includes(k.toLowerCase()) || k.toLowerCase().includes('nama')),
          status: Object.keys(firstRow).find(k => k.toLowerCase() === 'status'),
          date: Object.keys(firstRow).find(k => k.toLowerCase() === 'date' || k.toLowerCase() === 'tanggal')
        };

        if (flatMapping.name && flatMapping.status && flatMapping.date) {
           console.log('Flat Format Detected');
           upsertPayload = jsonData.map(row => {
             const pic_name = String(row[flatMapping.name!] || '').trim();
             const status = String(row[flatMapping.status!] || '').trim().toUpperCase();
             let schedule_date = '';
             
             const rawDate = row[flatMapping.date!];
             if (typeof rawDate === 'number') {
               const d = XLSX.SSF.parse_date_code(rawDate);
               schedule_date = format(new Date(d.y, d.m - 1, d.d), 'yyyy-MM-dd');
             } else {
               const d = new Date(rawDate);
               if (!isNaN(d.getTime())) schedule_date = format(d, 'yyyy-MM-dd');
             }
             
             return { pic_name, status, schedule_date };
           }).filter(r => r.pic_name && r.status && r.schedule_date);

        } else {
          // 2. MATRIX HEURISTIC (existing robust logic)
          console.log('Matrix Format Fallback');
          let headerRowIndex = -1;
          let nameColumnIndex = -1;
          let dateHeaders: any[] = [];

          for (let i = 0; i < Math.min(rawJsonAoA.length, 100); i++) {
            const row = rawJsonAoA[i];
            if (!row || !Array.isArray(row)) continue;
            
            const picIdx = row.findIndex(c => {
              const s = String(c || '').toLowerCase().trim();
              return s.includes('nama') || s.includes('pic') || s === 'name' || s === 'karyawan';
            });
            
            if (picIdx !== -1) {
              headerRowIndex = i;
              nameColumnIndex = picIdx;
              dateHeaders = row.slice(nameColumnIndex + 1);
              break;
            }
          }

          if (headerRowIndex !== -1) {
            const normalizeDate = (raw: any): string | null => {
              if (!raw) return null;
              const str = String(raw).trim().toLowerCase();
              if (!isNaN(Number(str)) && Number(str) > 10000) {
                const d = XLSX.SSF.parse_date_code(Number(str));
                return format(new Date(d.y, d.m - 1, d.d), 'yyyy-MM-dd');
              }
              const monthMap: Record<string, string> = { 
                'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 'may': '05', 'mei': '05', 
                'jun': '06', 'jul': '07', 'aug': '08', 'agu': '08', 'ags': '08', 'sep': '09', 
                'oct': '10', 'okt': '10', 'nov': '11', 'dec': '12', 'des': '12' 
              };
              const match = str.match(/(\d{1,2})[\s\-\/]([a-z]{3}|\d{1,2})/);
              if (match) {
                const day = match[1].padStart(2, '0');
                const monthPart = match[2];
                let m = /^\d+$/.test(monthPart) ? monthPart.padStart(2, '0') : (monthMap[monthPart.substring(0, 3)] || '');
                if (m && day) return `${new Date().getFullYear()}-${m}-${day}`;
              }
              try {
                const d = new Date(str);
                if (!isNaN(d.getTime())) return format(d, 'yyyy-MM-dd');
              } catch (e) {}
              return null;
            };

            const parsedDates = dateHeaders.map(d => normalizeDate(d));
            const dataRows = rawJsonAoA.slice(headerRowIndex + 1);

            dataRows.forEach(row => {
              const picName = String(row[nameColumnIndex] || '').trim();
              if (!picName || picName.length < 2 || /^\d+$/.test(picName)) return;

              parsedDates.forEach((scheduleDate, dateIdx) => {
                if (!scheduleDate) return;
                const rawValue = String(row[nameColumnIndex + 1 + dateIdx] || '').trim().toUpperCase();
                let status = 'WFH';
                if (rawValue === 'WFO') status = 'WFO';
                else if (['LIBUR', 'OFF', 'CUTI'].includes(rawValue)) status = 'LIBUR';
                else if (['WFC', 'A2'].includes(rawValue)) status = 'WFC';
                else if (rawValue === 'WFH') status = 'WFH';
                
                upsertPayload.push({ pic_name: picName, status, schedule_date: scheduleDate });
              });
            });
          }
        }

        if (upsertPayload.length > 0) {
          const chunkSize = 100;
          for (let i = 0; i < upsertPayload.length; i += chunkSize) {
            await taskService.upsertSchedules(upsertPayload.slice(i, i + chunkSize));
          }
          alert(`Berhasil mengimpor ${upsertPayload.length} data jadwal.`);
          fetchSchedules();
        } else {
          alert('Format data tidak dikenali (Gunakan CSV Flat atau Matrix)');
        }
      } catch (err) {
        console.error('Import Error:', err);
        alert(err instanceof Error ? err.message : 'Gagal mengimpor file.');
      } finally {
        setIsImporting(false);
        if (e.target) e.target.value = '';
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleStatusUpdate = async (pic: string, date: string, status: string) => {
    const currentStatus = scheduleGrid[pic]?.[date] || '';
    
    // Ownership check: Only own row or Admin/Superadmin
    const isOwner = fullName === pic;
    if (!isOwner && !isAdmin) {
      alert('Bukan jadwal Anda. Anda hanya dapat mengajukan perubahan untuk jadwal atas nama Anda.');
      setEditingCell(null);
      return;
    }

    // Anti-Spam: Prevent Duplicate Pending Requests
    try {
      const isPending = await taskService.checkExistingRescheduleRequest(pic, date);
      if (isPending && !isAdmin) {
        alert("Jadwal ini sedang dalam status Pending approval. Silakan tunggu response Admin.");
        setEditingCell(null);
        return;
      }
    } catch (err) {
      console.error("Duplicate check failed:", err);
    }

    // Forced Approval Workflow (Test Mode): Every change must go to requests
    setRequestModal({ pic, date, currentStatus, newStatus: status });
    setEditingCell(null);
  };

  const handleDownloadTemplate = () => {
    const data = [
      ["Name", "02 Jan 2026", "03 Jan 2026", "04 Jan 2026", "05 Jan 2026"],
      ["Fachrul Wisnu Novianto", "WFO", "WFH", "LIBUR", ""]
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template_Matrix");
    XLSX.writeFile(wb, "Template_Schedule_Sakti_Matrix.xlsx");
  };

  const handleExport = () => {
    // Generate grid for Excel
    const header = ['PIC', ...days.map(d => format(d, 'dd/MM/yyyy'))];
    const dataRows = pics.map(pic => [
      pic,
      ...days.map(d => scheduleGrid[pic]?.[format(d, 'yyyy-MM-dd')] || '-')
    ]);

    const ws = XLSX.utils.aoa_to_sheet([header, ...dataRows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Schedule");
    XLSX.writeFile(wb, `OMDEDY_Schedule_${format(currentMonth, 'MMM_yyyy')}.xlsx`);
  };

  return (
    <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-700">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
        <div className="flex items-center gap-4">
          <div className="w-14 h-14 bg-indigo-500/10 rounded-2xl flex items-center justify-center border border-indigo-500/20 shadow-lg shadow-indigo-500/5">
            <Calendar className="w-7 h-7 text-indigo-400" />
          </div>
          <div>
            <h2 className="text-2xl font-black text-white tracking-tighter italic uppercase">Om Dedy Schedule</h2>
            <p className="text-xs text-slate-500 font-bold uppercase tracking-[0.2em] leading-none mt-1">Resource Capacity Monitoring</p>
          </div>
        </div>

        <div className="flex items-center gap-4 bg-slate-900/80 p-1.5 rounded-2xl border border-slate-800 shadow-xl">
           <button onClick={() => setCurrentMonth(addMonths(currentMonth, -1))} className="p-2.5 hover:bg-slate-800 rounded-xl text-slate-400 transition-all hover:text-white">
             <ChevronLeft className="w-5 h-5" />
           </button>
           <div className="px-6 py-1 border-x border-slate-800">
             <span className="text-lg font-black text-white italic tracking-widest uppercase">{format(currentMonth, 'MMMM yyyy')}</span>
           </div>
           <button onClick={() => setCurrentMonth(addMonths(currentMonth, 1))} className="p-2.5 hover:bg-slate-800 rounded-xl text-slate-400 transition-all hover:text-white">
             <ChevronRight className="w-5 h-5" />
           </button>
        </div>

        <div className="flex items-center gap-3">
          {(user.access_level === 'Superadmin' || user.access_level === 'Admin') && (
            <>
              <button 
                onClick={() => setShowBatchModal(true)}
                className="flex items-center gap-2 px-5 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95 border border-slate-700"
              >
                <Plus className="w-4 h-4" />
                Input Massal
              </button>
              <button 
                onClick={() => setActiveView('RESCHEDULE')}
                className="flex items-center gap-2 px-5 py-2.5 bg-amber-500/10 hover:bg-amber-500/20 text-amber-500 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95 border border-amber-500/20 relative"
              >
                <History className="w-4 h-4" />
                Reschedule
                {rescheduleRequests.length > 0 && (
                  <span className="absolute -top-1 -right-1 w-4 h-4 bg-rose-500 text-white text-[8px] rounded-full flex items-center justify-center animate-bounce shadow-lg ring-2 ring-slate-950">
                    {rescheduleRequests.length}
                  </span>
                )}
              </button>
            </>
          )}
          <label className="flex items-center gap-2 px-5 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-[10px] font-black uppercase tracking-widest cursor-pointer shadow-lg shadow-indigo-500/20 transition-all active:scale-95 group">
            <Upload className="w-4 h-4 group-hover:bounce" />
            {isImporting ? 'Parsing...' : 'Smart Import'}
            <input type="file" className="hidden" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} disabled={isImporting} />
          </label>
          <button onClick={handleExport} className="flex items-center gap-2 px-5 py-2.5 bg-black/20 hover:bg-black/30 text-slate-400 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95 border border-white/5">
            <Download className="w-4 h-4" />
            Export Grid
          </button>
          <button 
            onClick={handleDownloadTemplate}
            className="flex items-center gap-2 px-5 py-2.5 bg-slate-900 border border-slate-800 hover:bg-slate-800 text-slate-400 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95 shadow-xl"
          >
            <History className="w-4 h-4" />
            Template Excel
          </button>
        </div>

        {showBatchModal && (
          <BatchManualEntryModal 
            users={users} 
            onClose={() => setShowBatchModal(false)} 
            onSuccess={() => {
              fetchSchedules();
            }} 
          />
        )}
      </div>

      {/* Monthly Statistics Summary */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <div className="bg-slate-900/50 border border-slate-800 rounded-3xl p-5 flex items-center gap-4 group hover:bg-emerald-500/5 transition-all">
          <div className="w-12 h-12 bg-emerald-500/10 rounded-2xl flex items-center justify-center border border-emerald-500/20 shadow-lg shadow-emerald-500/5">
            < Rocket className="w-6 h-6 text-emerald-400" />
          </div>
          <div>
            <p className="text-[9px] font-black text-slate-500 uppercase tracking-[0.2em] mb-1">Total WFO</p>
            <h4 className="text-2xl font-black text-white italic">{stats.WFO}</h4>
          </div>
        </div>
        <div className="bg-slate-900/50 border border-slate-800 rounded-3xl p-5 flex items-center gap-4 group hover:bg-indigo-500/5 transition-all">
          <div className="w-12 h-12 bg-indigo-500/10 rounded-2xl flex items-center justify-center border border-indigo-500/20 shadow-lg shadow-indigo-500/5">
            <Activity className="w-6 h-6 text-indigo-400" />
          </div>
          <div>
            <p className="text-[9px] font-black text-slate-500 uppercase tracking-[0.2em] mb-1">Total WFH</p>
            <h4 className="text-2xl font-black text-white italic">{stats.WFH}</h4>
          </div>
        </div>
        <div className="bg-slate-900/50 border border-slate-800 rounded-3xl p-5 flex items-center gap-4 group hover:bg-rose-500/5 transition-all">
          <div className="w-12 h-12 bg-rose-500/10 rounded-2xl flex items-center justify-center border border-rose-500/20 shadow-lg shadow-rose-500/5">
            <Filter className="w-6 h-6 text-rose-400" />
          </div>
          <div>
            <p className="text-[9px] font-black text-slate-500 uppercase tracking-[0.2em] mb-1">Absence/Libur</p>
            <h4 className="text-2xl font-black text-white italic">{stats.CUTI_LIBUR}</h4>
          </div>
        </div>
        <div className="bg-slate-900/50 border border-slate-800 rounded-3xl p-5 flex items-center gap-4 group hover:bg-amber-500/5 transition-all">
          <div className="w-12 h-12 bg-amber-500/10 rounded-2xl flex items-center justify-center border border-amber-500/20 shadow-lg shadow-amber-500/5">
            <LayoutGrid className="w-6 h-6 text-amber-400" />
          </div>
          <div>
            <p className="text-[9px] font-black text-slate-500 uppercase tracking-[0.2em] mb-1">Outstation/WFC</p>
            <h4 className="text-2xl font-black text-white italic">{stats.WFC}</h4>
          </div>
        </div>
      </div>

      <div className="bg-slate-900/30 border border-slate-800/80 rounded-[2.5rem] overflow-hidden flex flex-col backdrop-blur-xl shadow-2xl">
        <div className="overflow-auto relative min-h-[500px]">
          <table className="w-full border-collapse">
            <thead>
              <tr className="bg-slate-900/90 border-b border-slate-800 sticky top-0 z-40 backdrop-blur-md">
                <th className="sticky left-0 z-50 bg-slate-950 px-8 py-5 border-r border-slate-800 min-w-[240px] text-left">
                  <div className="flex items-center gap-3">
                    <div className="w-2 h-2 rounded-full bg-indigo-500 animate-pulse" />
                    <span className="text-[10px] font-black text-slate-500 uppercase tracking-[0.3em]">PIC RESOURCE</span>
                  </div>
                </th>
                {days.map((d, di) => (
                  <th key={`omdedy-header-${d.toISOString()}-${di}`} className={cn(
                    "px-4 py-5 border-r border-slate-800/20 min-w-[70px] transition-all",
                    isToday(d) ? "bg-indigo-500/20 ring-1 ring-inset ring-indigo-500/30" : "hover:bg-white/[0.02]"
                  )}>
                    <div className="flex flex-col gap-1 items-center relative group/header">
                       <span className={cn(
                         "text-xl font-black italic tracking-tighter",
                         isToday(d) ? "text-indigo-400" : (HOLIDAYS_2026[format(d, 'yyyy-MM-dd')] ? "text-rose-500" : (isWeekend(d) ? "text-rose-500/80" : "text-white"))
                       )}>
                         {format(d, 'dd')}
                       </span>
                       <span className={cn(
                         "text-[8px] font-black text-slate-500 uppercase tracking-widest",
                         CUTI_BERSAMA_2026[format(d, 'yyyy-MM-dd')] && "text-amber-500"
                       )}>
                         {format(d, 'EEE')}
                       </span>
                       
                       {(HOLIDAYS_2026[format(d, 'yyyy-MM-dd')] || CUTI_BERSAMA_2026[format(d, 'yyyy-MM-dd')]) && (
                         <div className="absolute top-full mt-2 hidden group-hover/header:block z-[100] bg-slate-800 border border-slate-700 p-2 rounded shadow-xl text-[10px] font-bold text-white whitespace-nowrap pointer-events-none">
                           {HOLIDAYS_2026[format(d, 'yyyy-MM-dd')] || CUTI_BERSAMA_2026[format(d, 'yyyy-MM-dd')]}
                         </div>
                       )}
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr>
                  <td colSpan={days.length + 1} className="py-32">
                     <div className="flex flex-col items-center gap-4">
                        <Activity className="w-10 h-10 text-indigo-500 animate-pulse" />
                        <p className="text-[10px] font-black text-slate-500 uppercase tracking-[0.4em]">Synchronizing Capacity Grid...</p>
                     </div>
                  </td>
                </tr>
              ) : pics.map(pic => (
                <tr key={pic} className="group border-b border-white/[0.02] hover:bg-white/[0.01] transition-all">
                   <td className="sticky left-0 z-30 bg-slate-950/95 backdrop-blur-md px-6 py-4 border-r border-slate-800 transition-colors">
                     <div className="flex flex-col gap-1.5">
                       <span className="font-black text-sm italic text-slate-300 tracking-tighter group-hover:text-indigo-400">
                         {pic}
                       </span>
                       <div className="flex items-center gap-2">
                         <div className="flex items-center gap-1 bg-emerald-500/10 px-1.5 py-0.5 rounded-md border border-emerald-500/20">
                           <span className="text-[8px] font-black text-emerald-500">O: {picStats[pic]?.WFO || 0}</span>
                         </div>
                         <div className="flex items-center gap-1 bg-indigo-500/10 px-1.5 py-0.5 rounded-md border border-indigo-500/20">
                           <span className="text-[8px] font-black text-indigo-400">H: {picStats[pic]?.WFH || 0}</span>
                         </div>
                         <div className="flex items-center gap-1 bg-slate-800 px-1.5 py-0.5 rounded-md border border-slate-700">
                           <span className="text-[8px] font-black text-slate-400">T: {picStats[pic]?.TOTAL || 0}</span>
                         </div>
                       </div>
                     </div>
                   </td>
                   {days.map((d, di) => {
                     const dateKey = format(d, 'yyyy-MM-dd');
                     const holiday = HOLIDAYS_2026[dateKey];
                     const cuti = CUTI_BERSAMA_2026[dateKey];
                     const status = scheduleGrid[pic]?.[dateKey];
                     
                     const displayStatus = status || (holiday ? 'LIBUR' : null);
                     
                     return (
                       <td key={`omdedy-cell-${pic}-${d.toISOString()}-${di}`} className={cn(
                         "p-2 border-r border-slate-800/10 text-center relative group/cell",
                         isToday(d) && "bg-indigo-500/5",
                         holiday && "bg-rose-500/5",
                         cuti && "bg-amber-500/5"
                       )}>
                         <div className="flex justify-center items-center h-full min-h-[40px]">
                            {(holiday || cuti) && (
                               <div className="absolute bottom-full mb-2 hidden group-hover/cell:block z-[100] bg-slate-800 border border-slate-700 p-2 rounded shadow-xl text-[10px] font-bold text-white whitespace-nowrap pointer-events-none">
                                 {holiday || cuti}
                               </div>
                            )}
                            
                            {editingCell?.pic === pic && editingCell?.date === dateKey ? (
                              <div className="absolute inset-x-1 z-[110] bg-slate-950/95 backdrop-blur-md p-2 flex flex-col gap-1.5 shadow-2xl border border-indigo-500/50 rounded-2xl animate-in zoom-in-95 duration-200">
                                 {['WFO', 'WFH', 'WFC', 'LIBUR', 'A2', 'DELETE'].map(st => (
                                   <button 
                                     key={st}
                                     onClick={() => {
                                       if (st === 'DELETE') {
                                          handleStatusUpdate(pic, dateKey, '');
                                       } else {
                                         handleStatusUpdate(pic, dateKey, st);
                                       }
                                     }}
                                     className={cn(
                                       "text-[9px] font-black px-3 py-1.5 rounded-xl grow hover:scale-105 active:scale-95 transition-all text-center border uppercase tracking-widest",
                                       st === 'DELETE' ? "bg-rose-500/20 text-rose-500 border-rose-500/30" : (SCHEDULE_STATUS_COLORS[st] || "bg-slate-800 text-slate-400 border-slate-700")
                                     )}
                                   >
                                     {st}
                                   </button>
                                 ))}
                                 <button onClick={() => setEditingCell(null)} className="text-[8px] font-black text-slate-500 bg-black/40 py-1.5 rounded-xl hover:text-white transition-colors">BATAL</button>
                              </div>
                            ) : (
                              displayStatus ? (
                               <motion.div 
                                 whileHover={{ scale: 1.1, zIndex: 10 }}
                                 onClick={() => {
                                   const isOwner = fullName === pic;
                                   if (isOwner || isAdmin) {
                                     setEditingCell({ pic, date: dateKey });
                                   } else {
                                     alert(`Bukan jadwal Anda. (PIC: ${pic}, Anda: ${fullName || 'Guest'})`);
                                   }
                                 }}
                                 className={cn(
                                   "px-2.5 py-1.5 rounded-xl text-[10px] font-black border text-center shadow-lg transition-all select-none",
                                   (fullName === pic || isAdmin) ? "cursor-pointer hover:ring-2 hover:ring-indigo-500/50" : "cursor-help",
                                   holiday ? "text-rose-100 bg-rose-600 border-rose-500" : (SCHEDULE_STATUS_COLORS[displayStatus] || 'text-slate-500 bg-slate-800/10 border-slate-800')
                                 )}
                               >
                                 {displayStatus}
                               </motion.div>
                             ) : (
                               <div 
                                 onClick={() => {
                                   const isOwner = fullName === pic;
                                   if (isOwner || isAdmin) {
                                     setEditingCell({ pic, date: dateKey });
                                   } else {
                                     alert(`Bukan jadwal Anda. (PIC: ${pic}, Anda: ${fullName || 'Guest'})`);
                                   }
                                 }}
                                 className={cn(
                                   "w-3 h-3 rounded-full bg-slate-800/50 transition-all shadow-inner",
                                   (fullName === pic || isAdmin) ? "cursor-pointer hover:bg-slate-600 hover:scale-125" : ""
                                 )} 
                               />
                             )
                            )}
                         </div>
                       </td>
                     );
                   })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Legend and Modals */}
      {requestModal && (
        <div className="fixed inset-0 z-[200] bg-slate-950/80 backdrop-blur-md flex items-center justify-center p-4">
           <motion.div 
             initial={{ opacity: 0, scale: 0.95, y: 20 }}
             animate={{ opacity: 1, scale: 1, y: 0 }}
             className="bg-slate-900 border border-white/10 rounded-[3rem] w-full max-w-lg shadow-[0_0_50px_rgba(0,0,0,0.5)] p-10 relative overflow-hidden"
           >
              {/* Glass Decorations */}
              <div className="absolute top-0 right-0 w-32 h-32 bg-indigo-500/10 rounded-full blur-3xl -translate-y-1/2 translate-x-1/2" />
              <div className="absolute bottom-0 left-0 w-32 h-32 bg-amber-500/5 rounded-full blur-3xl translate-y-1/2 -translate-x-1/2" />

              <div className="relative">
                <div className="flex items-center gap-5 mb-8">
                  <div className="w-14 h-14 bg-indigo-500/10 rounded-2xl flex items-center justify-center border border-indigo-500/20 shadow-lg shadow-indigo-500/5">
                    <Clock className="w-7 h-7 text-indigo-400" />
                  </div>
                  <div>
                    <h3 className="text-2xl font-black text-white italic uppercase tracking-tighter">Konfirmasi Perubahan Jadwal</h3>
                    <p className="text-[10px] font-bold text-slate-500 uppercase tracking-[0.3em] mt-1">Status: Pending Approval</p>
                  </div>
                </div>
                <div className="space-y-5 mb-10">
                  <div className="p-6 bg-slate-950/40 rounded-3xl border border-white/5 backdrop-blur-sm">
                    <div className="flex justify-between items-center mb-4">
                      <div>
                        <p className="text-[10px] text-slate-500 font-black uppercase tracking-widest mb-1">PIC TARGET</p>
                        <p className="text-md font-black text-white italic">{requestModal.pic}</p>
                      </div>
                      <div className="text-right">
                        <p className="text-[10px] text-slate-500 font-black uppercase tracking-widest mb-1">TANGGAL</p>
                        <p className="text-md font-black text-slate-300 italic">{format(new Date(requestModal.date), 'dd MMMM yyyy')}</p>
                      </div>
                    </div>

                    <div className="flex items-center gap-6 p-4 bg-black/40 rounded-2xl border border-white/5 relative overflow-hidden group">
                      <div className="flex-1">
                        <p className="text-[8px] text-slate-600 font-black uppercase tracking-widest mb-1">Original</p>
                        <span className="text-xs font-black text-slate-500">{requestModal.currentStatus || 'KOSONG'}</span>
                      </div>
                      <div className="w-8 h-8 bg-indigo-500/10 rounded-full flex items-center justify-center border border-indigo-500/20">
                        <ChevronRight className="w-4 h-4 text-indigo-400" />
                      </div>
                      <div className="flex-1 text-right">
                        <p className="text-[8px] text-indigo-500 font-black uppercase tracking-widest mb-1">Proposed</p>
                        <span className="text-xs font-black text-indigo-400">{requestModal.newStatus || 'DELETE'}</span>
                      </div>
                    </div>
                  </div>

                  {/* Two-Way Swap Feature */}
                  <div className="p-6 bg-slate-950/20 rounded-3xl border border-white/5">
                    <label className="flex items-center gap-3 cursor-pointer group mb-4">
                      <input 
                        type="checkbox" 
                        checked={requestModal.isSwap || false}
                        onChange={(e) => setRequestModal({ ...requestModal, isSwap: e.target.checked })}
                        className="w-5 h-5 rounded-lg border-slate-700 bg-slate-950 text-indigo-600 focus:ring-indigo-500/50"
                      />
                      <span className="text-xs font-black text-slate-300 uppercase tracking-widest group-hover:text-white transition-colors">Tukar 2 Arah (Opsional)</span>
                    </label>

                    {requestModal.isSwap && (
                      <motion.div 
                        initial={{ opacity: 0, height: 0 }}
                        animate={{ opacity: 1, height: 'auto' }}
                        className="space-y-4 pt-2 overflow-hidden"
                      >
                        <div className="grid grid-cols-2 gap-4">
                          <div className="space-y-2">
                            <label className="text-[8px] text-slate-500 font-black uppercase tracking-widest ml-1">Tanggal Pengganti</label>
                            <input 
                              type="date"
                              value={requestModal.swapDate || ''}
                              onChange={(e) => {
                                const date = e.target.value;
                                const swapCurrentStatus = scheduleGrid[requestModal.pic]?.[date] || '';
                                setRequestModal({ ...requestModal, swapDate: date, swapCurrentStatus });
                              }}
                              className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-indigo-500/50"
                            />
                          </div>
                          <div className="space-y-2">
                            <label className="text-[8px] text-slate-500 font-black uppercase tracking-widest ml-1">Status Pengganti</label>
                            <select 
                              value={requestModal.swapStatus || ''}
                              onChange={(e) => setRequestModal({ ...requestModal, swapStatus: e.target.value })}
                              className="w-full bg-slate-950 border border-white/10 rounded-xl px-4 py-2.5 text-xs text-indigo-400 font-black uppercase focus:outline-none focus:ring-1 focus:ring-indigo-500/50"
                            >
                              <option value="">Pilih Status</option>
                              {['WFO', 'WFH', 'A2', 'LIBUR'].map(st => (
                                <option key={st} value={st}>{st}</option>
                              ))}
                            </select>
                          </div>
                        </div>
                        {requestModal.swapDate && (
                          <div className="p-3 bg-black/20 rounded-xl border border-white/5 text-center">
                            <p className="text-[8px] text-slate-600 font-black uppercase tracking-tighter mb-1">Status Saat Ini pada {requestModal.swapDate}</p>
                            <span className="text-[10px] font-black text-slate-400">{requestModal.swapCurrentStatus || 'KOSONG'}</span>
                          </div>
                        )}
                      </motion.div>
                    )}
                  </div>

                  <div className="space-y-3">
                    <label className="text-[10px] text-slate-500 font-black uppercase tracking-widest ml-2">Alasan Reschedule</label>
                    <textarea 
                      id="reschedule-reason"
                      className="w-full bg-slate-950/50 border border-white/10 rounded-3xl p-5 text-sm text-white focus:outline-none focus:ring-2 focus:ring-indigo-500/30 transition-all placeholder:text-slate-700 min-h-[120px]"
                      placeholder="Contoh: Sakit, Ada keperluan mendesak, Tukar shift..."
                      rows={3}
                    />
                  </div>
                </div>

                <div className="flex gap-4">
                  <button 
                    onClick={() => setRequestModal(null)}
                    className="flex-1 py-4 bg-slate-800/50 hover:bg-slate-800 text-slate-400 rounded-2xl text-[11px] font-black uppercase tracking-[0.2em] transition-all border border-white/5"
                  >
                    Batalkan
                  </button>
                  <button 
                    onClick={async () => {
                      const reason = (document.getElementById('reschedule-reason') as HTMLTextAreaElement).value;
                      if (!reason.trim()) {
                        alert('Harap masukkan alasan penukaran jadwal.');
                        return;
                      }

                      if (requestModal.isSwap) {
                        if (!requestModal.swapDate || !requestModal.swapStatus) {
                          alert('Harap lengkapi detail tukar jadwal.');
                          return;
                        }
                        if (requestModal.swapDate === requestModal.date) {
                          alert('Tanggal pengganti tidak boleh sama dengan tanggal utama.');
                          return;
                        }

                        // Check duplicate for swap date
                        try {
                          const isSwapPending = await taskService.checkExistingRescheduleRequest(requestModal.pic, requestModal.swapDate);
                          if (isSwapPending) {
                            alert(`Tanggal pengganti (${requestModal.swapDate}) sedang dalam status Pending approval.`);
                            return;
                          }
                        } catch (e) { console.error(e); }
                      }

                      try {
                        setLoading(true);
                        
                        // Single payload for multi-way swaps
                        const payload = {
                          pic_name: requestModal.pic,
                          schedule_date: requestModal.date,
                          original_status: requestModal.currentStatus,
                          new_status: requestModal.newStatus,
                          reason: reason.trim() + (requestModal.isSwap ? ` (Tukar dengan ${requestModal.swapDate})` : ''),
                          requested_by: fullName || user.email || user.name,
                          swap_date: requestModal.isSwap ? requestModal.swapDate : null,
                          swap_status: requestModal.isSwap ? requestModal.swapStatus : null
                        };

                        await taskService.createRescheduleRequest(payload);
                        alert('Permohonan telah dikirim. Menunggu persetujuan Admin.');
                        setRequestModal(null);
                        fetchSchedules();
                      } catch (err: any) {
                        console.error(err);
                        alert(`Gagal mengirim permohonan: ${err.message || 'Unknown Error'}`);
                      } finally {
                        setLoading(false);
                      }
                    }}
                    className="flex-3 py-4 bg-indigo-600 hover:bg-indigo-500 text-white rounded-2xl text-[11px] font-black uppercase tracking-[0.2em] shadow-xl shadow-indigo-500/20 transition-all border border-white/20 active:scale-95"
                  >
                    Kirim Pengajuan
                  </button>
                </div>
              </div>
           </motion.div>
        </div>
      )}

      {showRescheduleRequests && (
        <div className="fixed inset-0 z-[200] bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4">
           <motion.div 
             initial={{ opacity: 0, scale: 0.95, y: 20 }}
             animate={{ opacity: 1, scale: 1, y: 0 }}
             className="bg-slate-900 border border-slate-800 rounded-[2.5rem] w-full max-w-4xl max-h-[85vh] shadow-2xl flex flex-col relative overflow-hidden"
           >
              <div className="p-8 border-b border-slate-800 flex justify-between items-center">
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 bg-amber-500/10 rounded-2xl flex items-center justify-center border border-amber-500/20">
                    <History className="w-6 h-6 text-amber-500" />
                  </div>
                  <div>
                    <h3 className="text-xl font-black text-white italic italic uppercase tracking-tight">Om Dedy: Reschedule Approval</h3>
                    <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Incoming and Pending Requests</p>
                  </div>
                </div>
                <button onClick={() => setShowRescheduleRequests(false)} className="p-3 bg-slate-800 hover:bg-slate-700 text-white rounded-2xl transition-all">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto p-8">
                {rescheduleRequests.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-20 gap-4 opacity-40">
                    <History className="w-16 h-16 text-slate-500" />
                    <p className="text-[10px] font-black text-slate-500 uppercase tracking-[0.4em]">No Pending Requests Found</p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    {rescheduleRequests.map((req: any) => (
                      <div key={req.id} className="bg-slate-950/50 border border-slate-800 p-6 rounded-3xl space-y-4 hover:border-indigo-500/30 transition-all">
                        <div className="flex justify-between items-start">
                          <div>
                            <p className="text-xs font-black text-white italic">{req.pic_name}</p>
                            <p className="text-[10px] text-slate-500 font-bold uppercase">{req.schedule_date}</p>
                          </div>
                          <div className={cn(
                            "px-3 py-1 text-[10px] font-black uppercase tracking-widest rounded-lg border",
                            req.status === 'Approved' ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/20" :
                            req.status === 'Rejected' ? "bg-rose-500/10 text-rose-500 border-rose-500/20" :
                            "bg-amber-500/10 text-amber-500 border-amber-500/20"
                          )}>
                            {req.status}
                          </div>
                        </div>
                        
                        <div className="flex items-center gap-4 bg-black/20 p-4 rounded-2xl">
                          <div className="flex-1 text-center">
                            <p className="text-[8px] text-slate-500 font-bold uppercase mb-1">From</p>
                            <span className="text-[10px] font-black text-slate-400 bg-slate-800 px-2.5 py-1 rounded-lg border border-slate-700">{req.original_status || 'BLANK'}</span>
                          </div>
                          <ChevronRight className="w-4 h-4 text-slate-700" />
                          <div className="flex-1 text-center">
                            <p className="text-[8px] text-indigo-500 font-bold uppercase mb-1">To</p>
                            <span className="text-[10px] font-black text-indigo-400 bg-indigo-500/10 px-2.5 py-1 rounded-lg border border-indigo-500/20">{req.new_status}</span>
                          </div>
                        </div>

                        <div className="p-4 bg-slate-900 rounded-2xl border border-slate-800/50">
                          <p className="text-[8px] text-slate-500 font-bold uppercase mb-2">Reason</p>
                          <p className="text-xs text-slate-300 italic">"{req.reason}"</p>
                        </div>
                        
                        {req.status === 'Pending' && (
                          <div className="flex gap-3">
                            <button 
                              onClick={async () => {
                                if (confirm('Tolak request ini?')) {
                                  try {
                                    await taskService.updateRescheduleRequestStatus(req.id, 'Rejected', user.email || 'Admin');
                                    fetchSchedules();
                                  } catch (err) { alert('Gagal memproses approval'); }
                                }
                              }}
                              className="flex-1 py-3 bg-slate-800 hover:bg-rose-500/10 hover:text-rose-400 text-slate-500 rounded-2xl text-[10px] font-black uppercase tracking-widest transition-all"
                            >
                              Reject
                            </button>
                            <button 
                              onClick={async () => {
                                if (confirm('Setujui request ini? Jadwal akan terupdate otomatis.')) {
                                  try {
                                    await taskService.updateRescheduleRequestStatus(req.id, 'Approved', user.email || 'Admin');
                                    fetchSchedules();
                                  } catch (err) { alert('Gagal memproses approval'); }
                                }
                              }}
                              className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-500 text-white rounded-2xl text-[10px] font-black uppercase tracking-widest transition-all shadow-lg shadow-emerald-500/20"
                            >
                              Approve
                            </button>
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>
           </motion.div>
        </div>
      )}

      <div className="bg-slate-900/50 border border-slate-800 p-6 rounded-[2rem] flex flex-wrap gap-8 items-center backdrop-blur-md">
         <div className="flex items-center gap-3 pr-8 border-r border-slate-800">
            <Filter className="w-4 h-4 text-slate-500" />
            <span className="text-[10px] font-black text-slate-500 uppercase tracking-[0.3em]">Status Legend</span>
         </div>
         {[
           { label: 'WFO - Work From Office', color: 'emerald', code: 'WFO' },
           { label: 'WFH - Work From Home', color: 'indigo', code: 'WFH' },
           { label: 'WFC/A2 - Outstation', color: 'amber', code: 'WFC' },
           { label: 'LIBUR/LATE - Absence', color: 'rose', code: 'LIBUR' },
         ].map(item => (
           <div key={item.code} className="flex items-center gap-3">
              <div className={cn(
                "w-4 h-4 rounded-lg border",
                SCHEDULE_STATUS_COLORS[item.code]
              )} />
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">{item.label}</span>
           </div>
         ))}
      </div>
    </div>
  );
}

function PersonelManagement({ users, projects, currentUser, onRefresh, isAdmin }: { users: AppUser[], projects: Project[], currentUser: any, onRefresh: () => void, isAdmin: boolean }) {
  const [searchTerm, setSearchTerm] = useState('');
  const [accessFilter, setAccessFilter] = useState('all');
  const [selectedPIC, setSelectedPIC] = useState<AppUser | null>(null);
  
  const [confirmConfig, setConfirmConfig] = useState<{ isOpen: boolean, title: string, description: string, onConfirm: () => void, variant?: 'danger' | 'primary' } | null>(null);
  const [showAddModal, setShowAddModal] = useState(false);
  const [showEditModal, setShowEditModal] = useState(false);
  const [editingUser, setEditingUser] = useState<AppUser | null>(null);
  const [notif, setNotif] = useState<string | null>(null);

  const filteredUsers = (users || []).filter(u => {
    const matchesSearch = u.name?.toLowerCase().includes(searchTerm.toLowerCase()) || u.email?.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesAccess = accessFilter === 'all' || u.access_level === accessFilter;
    return matchesSearch && matchesAccess;
  });

  const startEditing = (user: AppUser) => {
    setEditingUser(user);
    setShowEditModal(true);
  };

  const handleDelete = async (user: AppUser) => {
    if (!isAdmin) return;
    setConfirmConfig({
      isOpen: true,
      title: 'Hapus Personel?',
      description: `Apakah Anda yakin ingin menghapus ${user.name} (${user.email}) dari sistem? Tindakan ini tidak dapat dibatalkan.`,
      variant: 'danger',
      onConfirm: async () => {
        try {
          await taskService.deleteUser(user.id, currentUser?.email || 'System');
          setNotif(`Personnel ${user.name} removed successfully`);
          onRefresh();
          if (selectedPIC?.id === user.id) setSelectedPIC(null);
        } catch (err) {
          console.error('Delete user failed:', err);
        }
      }
    });
  };

  return (
    <div className="space-y-6">
      <SuccessNotification show={!!notif} message={notif || ''} onClose={() => setNotif(null)} />
      
      <AnimatePresence>
        {confirmConfig && (
          <ConfirmModal 
            isOpen={confirmConfig.isOpen}
            onClose={() => setConfirmConfig(null)}
            onConfirm={confirmConfig.onConfirm}
            title={confirmConfig.title}
            description={confirmConfig.description}
            variant={confirmConfig.variant}
            confirmText={confirmConfig.variant === 'primary' ? 'Ya, Simpan' : 'Hapus'}
          />
        )}
        {showAddModal && (
          <AddPersonnelModal 
            isAdmin={isAdmin}
            currentUserEmail={currentUser?.email || 'System'}
            onClose={() => setShowAddModal(false)}
            onSuccess={() => {
              setNotif('New personnel onboarded successfully');
              onRefresh();
            }}
          />
        )}
        {showEditModal && editingUser && (
          <EditPersonnelModal 
            user={editingUser}
            isAdmin={isAdmin}
            currentUserEmail={currentUser?.email || 'System'}
            onClose={() => {
              setShowEditModal(false);
              setEditingUser(null);
            }}
            onConfirmSave={async (updatedData) => {
              try {
                await taskService.updateUser(editingUser.id, updatedData, currentUser?.email || 'System');
                setNotif('Data berhasil diperbarui');
                setShowEditModal(false);
                setEditingUser(null);
                onRefresh();
              } catch (err: any) {
                alert('Update failed: ' + err.message);
              }
            }}
          />
        )}
      </AnimatePresence>

      <div className="flex flex-col md:flex-row gap-4 items-center justify-between mb-8">
        <div className="flex gap-4 w-full md:w-auto">
          <div className="relative flex-1 md:w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
            <input 
              type="text"
              placeholder="Search PIC or Email..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full bg-slate-900 border border-slate-800 rounded-xl pl-10 pr-4 py-2 text-xs text-slate-200 focus:border-indigo-500 outline-none transition-all"
            />
          </div>
          <select 
            value={accessFilter}
            onChange={(e) => setAccessFilter(e.target.value)}
            className="bg-slate-900 border border-slate-800 rounded-xl px-4 py-2 text-xs text-slate-400 focus:border-indigo-500 outline-none"
          >
            <option value="all">All Access Levels</option>
            <option value="Superadmin">Superadmin</option>
            <option value="Admin">Admin</option>
            <option value="PIC">PIC</option>
            <option value="Developer">Developer</option>
            <option value="QA">QA</option>
          </select>
        </div>
        {isAdmin && (
          <button 
            onClick={() => setShowAddModal(true)}
            className="bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest shadow-lg shadow-indigo-500/20 flex items-center gap-2 transition-all active:scale-95"
          >
            <UserPlus className="w-4 h-4" /> Add Personnel
          </button>
        )}
      </div>

      <div className="bg-slate-900/50 border border-slate-800/60 rounded-2xl overflow-hidden shadow-2xl">
        <table className="w-full text-left">
          <thead>
            <tr className="border-b border-slate-800 bg-slate-900/50">
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Nama PIC</th>
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Email PIC</th>
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Password</th>
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Access Level</th>
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Role</th>
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em] text-right">Action</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800/50">
            {filteredUsers.map((u, i) => {
              const canEdit = isAdmin || u.email === currentUser?.email;
              const userUniqueId = u.id || `user-fallback-${i}-${crypto.randomUUID()}`;

              return (
                <tr 
                  key={`personnel-${userUniqueId}`} 
                  className="hover:bg-indigo-500/5 transition-all group cursor-pointer border-b border-slate-800/30"
                  onClick={() => setSelectedPIC(u)}
                >
                  <td className="px-6 py-4">
                    <div className="flex items-center gap-3">
                      <div className="w-8 h-8 rounded-full bg-slate-800 border border-slate-700 flex items-center justify-center font-bold text-[10px] text-indigo-400 uppercase">
                        {u.name?.charAt(0)}
                      </div>
                      <span className="font-bold text-slate-200">{u.name}</span>
                    </div>
                  </td>
                  <td className="px-6 py-4">
                     <span className="text-[11px] text-slate-400 font-mono">{u.email}</span>
                  </td>
                  <td className="px-6 py-4">
                     <span className="text-[11px] text-slate-500 font-mono">••••••••</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="px-2 py-1 rounded bg-indigo-500/10 border border-indigo-500/20 text-indigo-400 text-[10px] font-black uppercase tracking-wider">
                      {u.access_level}
                    </span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="font-black text-slate-600 text-[10px] uppercase tracking-[0.15em]">{u.role}</span>
                  </td>
                  <td className="px-6 py-4 text-right">
                    <div className="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                      {canEdit && (
                        <button 
                          onClick={e => { e.stopPropagation(); startEditing(u); }}
                          className="p-2 hover:bg-slate-800 rounded-lg text-slate-500 hover:text-indigo-400"
                          title="Edit Personnel"
                        >
                          <Edit3 className="w-4 h-4" />
                        </button>
                      )}
                      <button 
                        onClick={e => { e.stopPropagation(); setSelectedPIC(u); }}
                        className="p-2 hover:bg-slate-800 rounded-lg text-slate-500 hover:text-indigo-400"
                        title="View Info"
                      >
                        <ExternalLink className="w-4 h-4" />
                      </button>
                      {isAdmin && (
                        <button 
                          onClick={e => { e.stopPropagation(); handleDelete(u); }}
                          className="p-2 hover:bg-rose-500/10 rounded-lg text-slate-500 hover:text-rose-500"
                          title="Hapus Personel"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <AnimatePresence>
        {selectedPIC && (
          <div className="fixed inset-0 z-[60] flex justify-end">
            <motion.div 
              initial={{ x: '100%' }}
              animate={{ x: 0 }}
              exit={{ x: '100%' }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="w-full max-w-md bg-slate-900 border-l border-slate-800 shadow-2xl flex flex-col"
            >
              <div className="p-6 border-b border-slate-800 flex items-center justify-between bg-slate-900/50 backdrop-blur-md">
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 rounded-2xl bg-indigo-600 flex items-center justify-center text-xl font-black text-white italic">
                    {selectedPIC.name?.charAt(0)}
                  </div>
                  <div>
                    <h3 className="font-black text-white uppercase italic tracking-tighter text-lg">{selectedPIC.name}</h3>
                    <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">{selectedPIC.email}</p>
                  </div>
                </div>
                <button 
                  onClick={() => setSelectedPIC(null)}
                  className="p-2 hover:bg-slate-800 rounded-xl transition-colors text-slate-500"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
              <div className="flex-1 overflow-auto p-6 space-y-6">
                <div>
                   <h4 className="text-[10px] font-black text-slate-600 uppercase tracking-[0.2em] mb-4">Credentials</h4>
                   <div className="space-y-4">
                      <div className="space-y-1.5">
                        <label className="text-[9px] font-bold text-slate-500 uppercase px-1">Crypto-Key (Password)</label>
                        <input 
                          type="password"
                          defaultValue={selectedPIC.password || ''}
                          className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-300 outline-none focus:border-indigo-500 transition-all font-mono"
                          placeholder="Set password..."
                        />
                      </div>
                   </div>
                </div>

                <div>
                   <h4 className="text-[10px] font-black text-slate-600 uppercase tracking-[0.2em] mb-4">Assigned Projects</h4>
                   <div className="space-y-3">
                      {projects.filter(p => true).map((p, i) => { 
                        const projectUniqueId = p.id || `pic-project-fallback-${i}`;
                        return (
                          <div key={`pic-project-${projectUniqueId}-${i}`} className="bg-slate-950/50 border border-slate-800 rounded-xl p-4 flex items-center justify-between hover:border-indigo-500/50 transition-all group">
                            <div>
                              <p className="text-xs font-bold text-slate-200 uppercase">{p.name}</p>
                              <p className="text-[9px] text-slate-500 font-bold uppercase tracking-tighter mt-1">Status: <span className="text-indigo-400">{p.status}</span></p>
                            </div>
                            <button className="p-2 bg-slate-800 rounded-lg text-slate-500 hover:text-white transition-opacity opacity-0 group-hover:opacity-100">
                              <ArrowRight className="w-4 h-4" />
                            </button>
                          </div>
                        );
                      })}
                   </div>
                </div>
              </div>
            </motion.div>
            <div className="absolute inset-0 bg-black/40 backdrop-blur-sm -z-10" onClick={() => setSelectedPIC(null)} />
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function KanbanView({ projects, tasks, onOpenGantt, onUpdateProject }: { projects: Project[], tasks: Task[], onOpenGantt: (id: string) => void, onUpdateProject: (id: string, updates: Partial<Project>) => void }) {
  const [picFilter, setPicFilter] = useState('all');
  const [projectSearch, setProjectSearch] = useState('');

  const columns = [
    { id: ProjectStatus.FSD_PROGRESS, label: 'FSD on Progress', color: 'indigo' },
    { id: ProjectStatus.FSD_REVIEW, label: 'FSD on Review', color: 'amber' },
    { id: ProjectStatus.SIT_PROGRESS, label: 'SIT on Progress', color: 'emerald' },
    { id: ProjectStatus.UAT_PROGRESS, label: 'UAT on Progress', color: 'cyan' },
    { id: ProjectStatus.PROJECT_LATE, label: 'Project Late', color: 'rose' },
  ];

  const safeProjects = projects || [];
  const safeTasks = tasks || [];

  const filteredProjects = safeProjects.filter(p => {
    const matchesSearch = p.name.toLowerCase().includes(projectSearch.toLowerCase());
    return matchesSearch;
  });

  const getProjectsByStatus = (status: string) => {
    return (filteredProjects || []).filter(p => p.status === status);
  };

  return (
    <div className="flex flex-col h-full space-y-6">
      <header className="sticky top-0 z-20 bg-[#020617]/80 backdrop-blur-md py-4 border-b border-slate-800/50 mb-6 flex items-center justify-between gap-4">
        <div className="flex items-center gap-4 flex-1 max-w-2xl">
          <div className="relative flex-1">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-500" />
            <input 
              type="text"
              placeholder="Filter Project Name..."
              value={projectSearch}
              onChange={(e) => setProjectSearch(e.target.value)}
              className="w-full bg-slate-900/50 border border-slate-800 rounded-xl pl-9 pr-4 py-2 text-[11px] text-slate-300 focus:border-indigo-500 outline-none"
            />
          </div>
          <select 
            value={picFilter}
            onChange={(e) => setPicFilter(e.target.value)}
            className="bg-slate-900/50 border border-slate-800 rounded-xl px-4 py-2 text-[11px] text-slate-400 outline-none"
          >
            <option value="all">All PICs</option>
            {/* Unique PICs from tasks or users */}
          </select>
        </div>
        <div className="flex items-center gap-2 px-3 py-1 bg-slate-900/50 border border-slate-800 rounded-lg">
          <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
          <span className="text-[9px] font-black text-slate-500 uppercase tracking-widest">Real-time Sync Active</span>
        </div>
      </header>

      <div className="flex-1 overflow-x-auto snap-x pb-6 scrollbar-thin scrollbar-thumb-slate-700 scrollbar-track-transparent">
        <div className="flex gap-6 h-full min-w-max pr-24">
          {columns.map((col, i) => {
            const colProjects = col.id === 'Project Late' 
              ? filteredProjects.filter(p => p.status === 'Project Late') // Or logic for late
              : getProjectsByStatus(col.id);

            return (
              <div key={`${col.id}-${i}`} className="flex-1 w-[350px] min-w-[350px] shrink-0 bg-slate-900/30 rounded-2xl border border-slate-800/40 p-4 flex flex-col gap-4 snap-center">
                <div className="flex items-center justify-between px-2">
                  <div className="flex items-center gap-2">
                    <div className={`w-2 h-2 rounded-full bg-${col.color}-500 shadow-[0_0_8px_rgba(var(--${col.color}-500),0.4)]`} />
                    <h3 className="text-[10px] font-black text-slate-400 uppercase tracking-[0.15em]">{col.label}</h3>
                  </div>
                  <span className="bg-slate-800 px-2 py-0.5 rounded text-[10px] font-mono text-slate-500">{colProjects.length}</span>
                </div>

                <div className="flex-1 overflow-y-auto space-y-4 scrollbar-hide">
                  {(colProjects || []).map((p, i) => {
                    const projectTasks = safeTasks.filter(t => t.project_id === p.id);
                    const doneTasks = projectTasks.filter(t => t.status === TaskStatus.DONE).length;
                    const progress = projectTasks.length > 0 ? (doneTasks / projectTasks.length) * 100 : 0;
                    const projectUniqueId = p.id || `kanban-project-fallback-${i}`;
                    
                    return (
                      <motion.div 
                        layoutId={projectUniqueId}
                        key={`kanban-project-${projectUniqueId}-${i}`}
                        onClick={() => onUpdateProject(p.id, { status: col.id as ProjectStatus })} // Simulation of Drag/Click to move
                        className="bg-slate-900 border border-slate-800 hover:border-indigo-500/50 rounded-xl p-4 shadow-xl cursor-pointer transition-all hover:bg-slate-800/40 group relative overflow-hidden"
                      >
                        <div className="relative z-10">
                          <h4 className="text-xs font-black text-white italic uppercase tracking-tighter mb-2">{p.name}</h4>
                          <div className="flex flex-col gap-1.5 mb-4">
                             <div className="flex items-center gap-2 text-[9px] font-bold text-slate-500 uppercase tracking-widest">
                               <UserIcon className="w-3 h-3 text-indigo-500" />
                               <span>PIC: <span className="text-slate-300">{p.pic_name || p.leader_email || 'Unassigned'}</span></span>
                             </div>
                             <div className="flex items-center gap-2 text-[9px] font-bold text-slate-500 uppercase tracking-widest">
                               <Clock className="w-3 h-3 text-slate-600" />
                               <span>Updated: <span className="font-mono">{p.updated_at ? format(new Date(p.updated_at), 'MM/dd') : 'N/A'}</span></span>
                             </div>
                          </div>
                          
                          <div className="space-y-1">
                            <div className="flex justify-between text-[9px] font-bold uppercase tracking-tighter italic">
                              <span className="text-slate-500">Infrastructure Health</span>
                              <span className="text-indigo-400">{Math.round(progress)}%</span>
                            </div>
                            <div className="h-1 bg-slate-800 rounded-full overflow-hidden">
                              <motion.div 
                                initial={{ width: 0 }}
                                animate={{ width: `${progress}%` }}
                                className="h-full bg-indigo-500 shadow-[0_0_10px_#6366f1]"
                              />
                            </div>
                          </div>

                          <div className="mt-4 pt-4 border-t border-slate-800 flex items-center justify-between">
                            <button 
                              onClick={(e) => { e.stopPropagation(); onOpenGantt(p.id); }}
                              className="px-3 py-1.5 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 rounded-lg text-[9px] font-black uppercase tracking-widest text-white transition-all shadow-lg shadow-indigo-600/10 active:scale-95"
                            >
                              Open Gantt Detail
                            </button>
                            <div className="w-6 h-6 rounded-full bg-slate-950 flex items-center justify-center font-black text-[8px] text-indigo-500 border border-slate-800">
                              {p.name.charAt(0)}
                            </div>
                          </div>
                        </div>
                        <div className="absolute top-0 right-0 w-16 h-16 bg-indigo-500/5 blur-2xl -mr-8 -mt-8" />
                      </motion.div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}


function GanttDetailView({ 
  user, users, projectId, setFocusedProjectId, projects, tasks, hierarchicalTasks, expandedRows, scale, 
  setRefreshKey, handleToggleExpand, handleUpdateTask, handleOpenAudit, handleDeleteTask, 
  setScale, setTasks, onReschedule 
}: any) {
  const [isAdding, setIsAdding] = useState(false);
  const [taskModalData, setTaskModalData] = useState<{ parentId: string | null } | null>(null);
  const [activeTab, setActiveTab] = useState<'TASKS' | 'AUDIT'>('TASKS');
  const [rescheduleLogs, setRescheduleLogs] = useState<ProjectRescheduleLog[]>([]);
  // If projectId is null, we show the high-level Global Gantt
  const isGlobalView = !projectId;
  
  const currentProject = (projects || []).find((p: any) => p.id === projectId);

  const projectStats = useMemo(() => {
    if (!projectId || !tasks) return { minStart: null, maxEnd: null, totalManHours: 0 };
    const pTasks = tasks.filter((t: any) => t.project_id === projectId);
    if (pTasks.length === 0) return { minStart: null, maxEnd: null, totalManHours: 0 };

    const validDates = pTasks.flatMap(t => [t.start_time, t.end_time]).filter(Boolean);
    const minStart = validDates.length > 0 ? new Date(Math.min(...validDates.map(d => new Date(d).getTime()))) : null;
    const maxEnd = validDates.length > 0 ? new Date(Math.max(...validDates.map(d => new Date(d).getTime()))) : null;
    
    const totalManHours = pTasks
      .filter((t: any) => !t.parent_id) // Level 1 only
      .reduce((sum, t) => sum + (Number(t.duration_hours) || 0), 0);

    return { 
      minStart, 
      maxEnd, 
      totalManHours 
    };
  }, [tasks, projectId]);

  const displayStart = projectStats.minStart ? format(projectStats.minStart, 'yyyy-MM-dd') : 'Belum di-set';
  const displayEnd = projectStats.maxEnd ? format(projectStats.maxEnd, 'yyyy-MM-dd') : 'Belum di-set';
  
  useEffect(() => {
    if (projectId) {
      taskService.getProjectRescheduleLogs(projectId).then(setRescheduleLogs);
    }
  }, [projectId]);

  
  // RBAC & Filter logic
  const isSuperadmin = user?.role === 'Superadmin' || user?.email?.includes('wisnu');
  
  const filteredTasks = useMemo(() => {
    if (isGlobalView) return tasks || []; 
    const pTasks = (tasks || []).filter((t: any) => t.project_id === projectId);
    if (isSuperadmin) return pTasks;
    return pTasks.filter((t: any) => t.assignee === (user?.name || user?.email));
  }, [tasks, projectId, user, isSuperadmin, isGlobalView]);

  // Recalculate hierarchy for filtered tasks
  const filteredHierarchy = useMemo(() => {
    if (isGlobalView) {
      // For global view, roots are PROJECTS
      const globalRoots = (projects || []).map(p => {
        const pTasks = (tasks || []).filter(t => t.project_id === p.id);
        const startTimes = pTasks.map(t => new Date(t.start_time).getTime());
        const endTimes = pTasks.map(t => new Date(t.end_time).getTime());
        
        const minStart = startTimes.length > 0 ? Math.min(...startTimes) : new Date(p.created_at).getTime();
        const maxEnd = endTimes.length > 0 ? Math.max(...endTimes) : new Date(p.created_at).getTime() + (24 * 3600000); // +1 day if no tasks

        return {
          id: p.id,
          title: p.name,
          isProject: true,
          status: p.status,
          start_time: new Date(minStart).toISOString(),
          end_time: new Date(maxEnd).toISOString(),
          assignee: p.pic_name || p.leader_email,
          duration_hours: Math.round((maxEnd - minStart) / 3600000),
          duration_minutes: 0
        };
      });
      
      const map = new Map();
      globalRoots.forEach(r => map.set(r.id, [])); // No children displayed in high-level Gantt generally, OR we could show phases
      
      return { roots: globalRoots, map };
    }

    const roots: any[] = [];
    const map = new Map();
    (filteredTasks || []).forEach((t: any) => map.set(t.id, { ...t, children: [] }));
    (filteredTasks || []).forEach((t: any) => {
      if (t.parent_id && map.has(t.parent_id)) {
        map.get(t.parent_id).children.push(map.get(t.id));
      } else if (!t.parent_id) {
        roots.push(map.get(t.id));
      }
    });
    return { roots, map };
  }, [filteredTasks, isGlobalView, projects, tasks]);

  const handleExportWBS = () => {
    if (!currentProject) return;
    
    // Convert hierarchical roots into flat list for excel
    const flatList: any[] = [];
    const recurse = (t: any, level: number) => {
      flatList.push({
        'WBS Level': `Level ${level}`,
        'Task Name': t.title || t.name || 'Untitled',
        'From Date': t.start_time ? format(new Date(t.start_time), 'yyyy-MM-dd') : 'N/A',
        'To Date': t.end_time ? format(new Date(t.end_time), 'yyyy-MM-dd') : 'N/A',
        'Man Hours': t.duration_hours || 0,
        'Status': t.status || 'Unknown'
      });
      if (t.children && t.children.length > 0) {
        t.children.forEach((c: any) => recurse(c, level + 1));
      }
    };

    filteredHierarchy.roots.forEach((r: any) => recurse(r, 1));
    handleExcelExport(flatList, `Project_WBS_${currentProject.name.replace(/\s+/g, '_')}`);
  };

  if (!filteredHierarchy || !filteredHierarchy.roots || (filteredHierarchy.roots.length === 0 && !isGlobalView)) {
    return (
      <div className="h-full flex items-center justify-center bg-[#020617] text-slate-500 font-bold uppercase tracking-widest flex-col gap-4">
        <div className="w-12 h-12 rounded-full border-2 border-slate-800 flex items-center justify-center opacity-50">
          <Clock className="w-6 h-6" />
        </div>
        Timeline Partition Empty
      </div>
    );
  }

  return (
    <div className="flex flex-col h-full bg-[#020617] overflow-hidden">
       <AnimatePresence mode="wait">
       <motion.div 
         key={projectId || 'global'}
         initial={{ opacity: 0, scale: 0.98 }}
         animate={{ opacity: 1, scale: 1 }}
         exit={{ opacity: 0, scale: 1.02 }}
         transition={{ duration: 0.3 }}
         className="flex flex-col h-full"
       >
          {/* TOP: Task Manager (Only in Detail View) */}
          {!isGlobalView && (
            <div className="h-[45%] flex flex-col bg-slate-950/20 border border-slate-800/60 rounded-2xl m-4 overflow-hidden shadow-2xl shrink-0">
              <div className="p-4 border-b border-slate-800/60 flex items-center justify-between bg-slate-900/40">
                <div className="flex items-center gap-6">
                  <div className="flex bg-slate-950 p-1 rounded-xl border border-white/5">
                    <button 
                      onClick={() => setActiveTab('TASKS')}
                      className={cn(
                        "px-4 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-widest transition-all",
                        activeTab === 'TASKS' ? "bg-indigo-600 text-white shadow-lg shadow-indigo-600/20" : "text-slate-500 hover:text-slate-300"
                      )}
                    >
                      Infrastructure Breakdown
                    </button>
                    <button 
                      onClick={() => setActiveTab('AUDIT')}
                      className={cn(
                        "px-4 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-widest transition-all",
                        activeTab === 'AUDIT' ? "bg-indigo-600 text-white shadow-lg shadow-indigo-600/20" : "text-slate-500 hover:text-slate-300"
                      )}
                    >
                      Audit Trail
                    </button>
                  </div>

                  {currentProject && (
                    <div className="flex items-center gap-4 pl-4 border-l border-white/5">
                      <button 
                         onClick={handleExportWBS}
                         className="bg-slate-800 hover:bg-slate-700 text-slate-300 px-3 py-1.5 rounded-lg text-[8px] font-black uppercase tracking-widest flex items-center gap-2 border border-slate-700 transition-all active:scale-95"
                      >
                        <Download className="w-3.5 h-3.5" /> Export WBS
                      </button>
                      <div className="flex flex-col">
                        <span className="text-[8px] font-bold text-slate-500 uppercase tracking-widest leading-none mb-1">
                          {currentProject.pic_name || currentProject.leader_email || 'Lead PIC'} 
                          <span className="ml-2 text-sm text-slate-300 normal-case font-medium">
                            | 📅 {displayStart} s/d {displayEnd} | ⏱️ {projectStats.totalManHours.toFixed(1)} Hours
                          </span>
                        </span>
                        <div className="flex items-center gap-3">
                           <div className="flex items-center gap-2 bg-slate-950 px-3 py-1 rounded-lg border border-white/5 shadow-inner">
                             <Calendar className="w-3 h-3 text-indigo-400" />
                             <span className="text-[10px] text-indigo-100 font-mono italic">
                               {displayStart} - {displayEnd}
                             </span>
                           </div>
                           <div className="flex items-center gap-1.5 bg-indigo-500/10 px-3 py-1 rounded-lg border border-indigo-500/20">
                             <Clock className="w-3 h-3 text-indigo-400" />
                             <span className="text-[10px] text-indigo-300 font-black tracking-tight italic">
                               {projectStats.totalManHours.toFixed(1)} hrs
                             </span>
                           </div>
                           <button 
                             onClick={() => onReschedule(currentProject)}
                             className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-indigo-400 rounded-lg transition-colors border border-white/5"
                             title="Reschedule Project Timeline"
                           >
                             <ArrowDown className="w-3.5 h-3.5 rotate-[-90deg]" />
                           </button>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
                
                {activeTab === 'TASKS' && (
                  <button 
                    onClick={() => setTaskModalData({ parentId: null })}
                    className="bg-emerald-600 hover:bg-emerald-500 text-white px-4 py-2 rounded-xl text-[9px] font-black uppercase tracking-widest shadow-lg shadow-emerald-500/20 shadow-lg shadow-indigo-500/20 transition-all active:scale-95 flex items-center gap-2"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    New Infrastructure Node
                  </button>
                )}
              </div>
              
              <div className="flex-1 overflow-auto scrollbar-hide bg-slate-950/40 relative">
                {activeTab === 'TASKS' ? (
                  <GanttTree 
                    user={user}
                    users={users}
                    roots={filteredHierarchy.roots} 
                    map={filteredHierarchy.map} 
                    tasks={filteredTasks}
                    projects={projects}
                    expandedRows={expandedRows}
                    onToggleExpand={handleToggleExpand}
                    onUpdateTask={handleUpdateTask}
                    onOpenAudit={handleOpenAudit}
                    onDeleteTask={handleDeleteTask}
                    onAddSubTask={(parentId: string) => {
                      setTaskModalData({ parentId });
                    }}
                  />
                ) : (
                  <div className="p-8">
                     <AuditLogTable logs={rescheduleLogs} />
                  </div>
                )}

                {taskModalData && (
                  <CreateTaskModal 
                    projectId={projectId!}
                    parentId={taskModalData.parentId}
                    onClose={() => setTaskModalData(null)}
                    onSuccess={() => setRefreshKey((prev: any) => prev + 1)}
                    user={user}
                  />
                )}
              </div>
            </div>
          )}

         {/* Alternative TOP for Global View: Simple Legend / Filter */}
         {isGlobalView && (
           <div className="p-6 bg-slate-950/20 border-b border-slate-800/40 flex items-center justify-between mx-4 mt-4 rounded-2xl">
              <div>
                <h3 className="text-white font-black uppercase italic tracking-tighter text-lg">Portfolio Timeline Overview</h3>
                <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mt-1">Cross-Project Resource Analysis</p>
              </div>
              <div className="flex gap-6">
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 bg-emerald-500/40 border border-emerald-500 rounded" />
                  <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Done</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 bg-indigo-500/40 border border-indigo-500 rounded" />
                  <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">In Progress</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 bg-amber-500/40 border border-amber-500 rounded" />
                  <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Review</span>
                </div>
              </div>
           </div>
         )}

         {/* BOTTOM: Gantt Visualizer */}
         <div className="flex-1 min-h-0 bg-[#020617] border-t border-slate-800/60 flex flex-col">
           <div className="p-4 flex items-center justify-between bg-slate-950/30">
             <div className="flex items-center gap-3">
               <LayoutGrid className="w-4 h-4 text-indigo-500" />
               <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">Temporal Visualizer (Hourly Precision)</span>
             </div>
             <div className="flex bg-slate-900 p-1 rounded-lg border border-slate-800 scale-90">
               {(['MONTH', 'WEEK', 'DAY', 'HOUR'] as ViewScale[]).map((s, si) => (
                 <button
                   key={`detail-scale-${s}-${si}`}
                   onClick={() => setScale(s)}
                   className={cn(
                     "px-3 py-1 rounded-md text-[10px] font-bold tracking-wider transition-all",
                     scale === s ? "bg-slate-800 text-indigo-400 shadow-sm" : "text-slate-500 hover:text-slate-300"
                   )}
                 >
                   {s}
                 </button>
               ))}
             </div>
           </div>
           <div className="flex-1 overflow-auto bg-[#020617] scrollbar-hide relative border-t border-slate-800/40">
             <div className="flex min-w-[3600px] items-start h-full">
                <GanttTimeline 
                  user={user}
                  scale={scale} 
                  tasks={filteredTasks} 
                  hierarchicalTasks={filteredHierarchy}
                  expandedRows={expandedRows}
                  setTasks={setTasks} 
                  projects={projects}
                  isGlobalView={isGlobalView}
                  onSetFocus={(id) => {
                    setFocusedProjectId(id);
                  }}
                />
             </div>
           </div>
         </div>
       </motion.div>
       </AnimatePresence>
    </div>
  );
}

// --- Anti-Stuttering Input Component ---
function LocalInput({ 
  value, 
  onChange, 
  className, 
  placeholder, 
  type = 'text',
  required = false,
  autoFocus = false
}: { 
  value: string, 
  onChange: (v: string) => void, 
  className?: string, 
  placeholder?: string,
  type?: string,
  required?: boolean,
  autoFocus?: boolean
}) {
  const [local, setLocal] = useState(value);
  useEffect(() => { setLocal(value); }, [value]);

  return (
    <input 
      type={type}
      required={required}
      autoFocus={autoFocus}
      value={local}
      onChange={(e) => setLocal(e.target.value)}
      onBlur={() => onChange(local)}
      onKeyDown={(e) => e.key === 'Enter' && e.currentTarget.blur()}
      className={className}
      placeholder={placeholder}
    />
  );
}

function AddPersonnelModal({ 
  onClose, 
  onSuccess, 
  isAdmin,
  currentUserEmail
}: { 
  onClose: () => void, 
  onSuccess: () => void, 
  isAdmin: boolean,
  currentUserEmail: string
}) {
  const [formData, setFormData] = useState({
    name: '',
    email: '',
    access_level: 'PIC',
    role: 'Staff',
    password: 'password123'
  });
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    try {
      await taskService.createUser(formData, currentUserEmail);
      onSuccess();
      onClose();
    } catch (err: any) {
      setError(err.message || 'Failed to add personnel');
    }
  };

  return (
    <div className="fixed inset-0 z-[120] flex items-center justify-center p-4">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
        className="absolute inset-0 bg-slate-950/90 backdrop-blur-md"
      />
      <motion.div 
        initial={{ opacity: 0, scale: 0.9, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 20 }}
        className="relative w-full max-w-md bg-slate-900 border border-slate-800 rounded-3xl shadow-2xl overflow-hidden"
      >
        <div className="p-8 border-b border-white/5 bg-slate-900/50">
          <h3 className="text-2xl font-black text-white italic tracking-tighter uppercase">Onboard Personnel</h3>
          <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mt-1">Register new PIC into the OD Ecosystem</p>
        </div>

        <form onSubmit={handleSubmit} className="p-8 space-y-5">
          {error && <div className="p-3 bg-rose-500/10 border border-rose-500/20 text-rose-400 text-[10px] font-bold uppercase rounded-xl tracking-wider">{error}</div>}
          
          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Nama PIC</label>
            <LocalInput 
              required
              value={formData.name}
              onChange={v => setFormData({...formData, name: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 transition-all font-bold"
              placeholder="e.g., John Doe"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Email PIC</label>
            <LocalInput 
              required
              type="email"
              value={formData.email}
              onChange={v => setFormData({...formData, email: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 transition-all font-mono"
              placeholder="email@example.com"
            />
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Access Level</label>
              <select 
                value={formData.access_level}
                onChange={e => setFormData({...formData, access_level: e.target.value})}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-indigo-400 outline-none focus:border-indigo-500 font-bold"
              >
                <option value="Superadmin">Superadmin</option>
                <option value="Admin">Admin</option>
                <option value="PIC">PIC</option>
                <option value="Developer">Developer</option>
                <option value="QA">QA</option>
              </select>
            </div>
            <div className="space-y-1.5">
              <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Role (Text)</label>
              <LocalInput 
                value={formData.role}
                onChange={v => setFormData({...formData, role: v})}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 font-bold"
                placeholder="Developer, QA, etc."
              />
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Security-Key</label>
            <LocalInput 
              required
              type="password"
              value={formData.password}
              onChange={v => setFormData({...formData, password: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 font-mono"
              placeholder="••••••••"
            />
          </div>

          <div className="pt-4 flex gap-4">
            <button 
              type="button"
              onClick={onClose}
              className="flex-1 py-3 bg-slate-800 hover:bg-slate-700 text-slate-300 text-[10px] font-black uppercase tracking-widest rounded-xl transition-all"
            >
              Cancel
            </button>
            <button 
              type="submit"
              className="flex-3 py-3 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white text-[10px] font-black uppercase tracking-widest rounded-xl shadow-lg shadow-indigo-500/20 transition-all"
            >
              Onboard User
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
}

function EditPersonnelModal({ 
  user,
  onClose, 
  onConfirmSave, 
  isAdmin,
  currentUserEmail
}: { 
  user: AppUser,
  onClose: () => void, 
  onConfirmSave: (data: Partial<AppUser>) => void, 
  isAdmin: boolean,
  currentUserEmail: string
}) {
  const [formData, setFormData] = useState({
    name: user.name || '',
    email: user.email || '',
    access_level: user.access_level || 'PIC',
    role: user.role || 'Staff',
    password: user.password || ''
  });

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onConfirmSave(formData);
  };

  return (
    <div className="fixed inset-0 z-[120] flex items-center justify-center p-4">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
        className="absolute inset-0 bg-slate-950/90 backdrop-blur-md"
      />
      <motion.div 
        initial={{ opacity: 0, scale: 0.9, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 20 }}
        className="relative w-full max-w-md bg-slate-900 border border-slate-800 rounded-3xl shadow-2xl overflow-hidden"
      >
        <div className="p-8 border-b border-white/5 bg-slate-900/50">
          <h3 className="text-2xl font-black text-white italic tracking-tighter uppercase">Edit Personnel</h3>
          <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mt-1">Modify credentials for {user.name}</p>
        </div>

        <form onSubmit={handleSubmit} className="p-8 space-y-5">
          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Nama PIC</label>
            <LocalInput 
              required
              value={formData.name}
              onChange={v => setFormData({...formData, name: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 transition-all font-bold"
              placeholder="Full Name"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Email PIC</label>
            <LocalInput 
              required
              type="email"
              value={formData.email}
              onChange={v => setFormData({...formData, email: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-400 outline-none focus:border-indigo-500 transition-all font-mono read-only:opacity-60"
              placeholder="email@example.com"
            />
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Access Level</label>
              <select 
                value={formData.access_level}
                disabled={!isAdmin}
                onChange={e => setFormData({...formData, access_level: e.target.value})}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-indigo-400 outline-none focus:border-indigo-500 font-bold disabled:opacity-60"
              >
                <option value="Superadmin">Superadmin</option>
                <option value="Admin">Admin</option>
                <option value="PIC">PIC</option>
                <option value="Developer">Developer</option>
                <option value="QA">QA</option>
              </select>
            </div>
            <div className="space-y-1.5">
              <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Role (Text)</label>
              <LocalInput 
                value={formData.role}
                onChange={v => setFormData({...formData, role: v})}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 font-bold"
                placeholder="Developer, QA, etc."
              />
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Security-Key</label>
            <LocalInput 
              type="password"
              value={formData.password}
              onChange={v => setFormData({...formData, password: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 font-mono"
              placeholder="••••••••"
            />
          </div>

          <div className="pt-4 flex gap-4">
            <button 
              type="button"
              onClick={onClose}
              className="flex-1 py-3 bg-slate-800 hover:bg-slate-700 text-slate-300 text-[10px] font-black uppercase tracking-widest rounded-xl transition-all"
            >
              Cancel
            </button>
            <button 
              type="submit"
              className="flex-3 py-3 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white text-[10px] font-black uppercase tracking-widest rounded-xl shadow-lg shadow-indigo-500/20 transition-all font-black uppercase tracking-widest shadow-lg shadow-indigo-500/20 flex items-center justify-center gap-2"
            >
              Update Data
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
}

function RescheduleRequestsView({ requests, onRefresh, user }: { requests: any[], onRefresh: () => void, user: any }) {
  const [confirmData, setConfirmData] = useState<{ id: string, status: 'Approved' | 'Rejected' } | null>(null);
  const [activeTab, setActiveTab] = useState<'PENDING' | 'HISTORY'>('PENDING');
  const [deleteId, setDeleteId] = useState<string | null>(null);

  if (!user || (user.access_level?.toLowerCase() !== 'superadmin' && user.access_level?.toLowerCase() !== 'admin')) {
    return (
      <div className="flex items-center justify-center h-full">
        <div className="text-center space-y-4">
          <ShieldAlert className="w-16 h-16 text-rose-500 mx-auto" />
          <h2 className="text-2xl font-black text-white uppercase italic">Unauthorized Access</h2>
          <p className="text-slate-500">You do not have permission to access this module (Reschedule Om Dedy).</p>
        </div>
      </div>
    );
  }

  const handleAction = async () => {
    if (!confirmData) return;
    try {
      await taskService.updateRescheduleRequestStatus(confirmData.id, confirmData.status, user.email || 'Admin');
      onRefresh();
    } catch (err) { 
      alert('Gagal memproses approval'); 
    } finally {
      setConfirmData(null);
    }
  };

  const handleDelete = async () => {
    if (!deleteId) return;
    try {
      await taskService.deleteRescheduleRequest(deleteId, user.email || 'Admin');
      onRefresh();
    } catch (err) {
      alert('Gagal menghapus request');
    } finally {
      setDeleteId(null);
    }
  };

  const pendingRequests = requests.filter(r => r.status === 'Pending');
  const historyRequests = requests.filter(r => r.status !== 'Pending');
  const currentRequests = activeTab === 'PENDING' ? pendingRequests : historyRequests;

  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-700 h-full flex flex-col">
      <ConfirmModal 
        isOpen={!!confirmData}
        onClose={() => setConfirmData(null)}
        onConfirm={handleAction}
        title={confirmData?.status === 'Approved' ? "Setujui Request?" : "Tolak Request?"}
        description={confirmData?.status === 'Approved' 
          ? "Jadwal personel akan diperbarui secara otomatis di sistem." 
          : "Permohonan reschedule ini akan ditolak dan personel akan diberitahu."}
        variant={confirmData?.status === 'Approved' ? 'primary' : 'danger'}
        confirmText={confirmData?.status === 'Approved' ? "Approve" : "Reject"}
      />

      <ConfirmModal 
        isOpen={!!deleteId}
        onClose={() => setDeleteId(null)}
        onConfirm={handleDelete}
        title="Hapus History Request?"
        description="Data ini akan dihapus permanen dari audit trail reschedule."
        variant="danger"
        confirmText="Hapus Permanen"
      />

      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <div className="w-14 h-14 bg-amber-500/10 rounded-2xl flex items-center justify-center border border-amber-500/20 shadow-lg shadow-amber-500/5">
            <History className="w-7 h-7 text-amber-400" />
          </div>
          <div>
            <h2 className="text-2xl font-black text-white tracking-tighter italic uppercase">Reschedule Om Dedy</h2>
            <p className="text-xs text-slate-500 font-bold uppercase tracking-[0.2em] mt-1">Personnel Swap Request Management</p>
          </div>
        </div>

        <div className="flex bg-slate-900/50 p-1 rounded-xl border border-white/5">
          <button 
            onClick={() => setActiveTab('PENDING')}
            className={cn(
              "px-6 py-2 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all flex items-center gap-2",
              activeTab === 'PENDING' ? "bg-amber-600 text-white shadow-lg shadow-amber-600/20" : "text-slate-500 hover:text-slate-300"
            )}
          >
            Pending Requests
            <span className="bg-black/40 px-1.5 py-0.5 rounded-md text-[8px]">{pendingRequests.length}</span>
          </button>
          <button 
            onClick={() => setActiveTab('HISTORY')}
            className={cn(
              "px-6 py-2 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all flex items-center gap-2",
              activeTab === 'HISTORY' ? "bg-slate-700 text-white" : "text-slate-500 hover:text-slate-300"
            )}
          >
            Resolution History
          </button>
        </div>
      </div>

      <div className="flex-1 overflow-y-auto pr-4 scrollbar-hide">
        {currentRequests.length === 0 ? (
          <div className="h-full bg-slate-900/50 border border-slate-800 rounded-[2.5rem] p-20 flex flex-col items-center justify-center gap-4">
            <History className="w-16 h-16 text-slate-700" />
            <p className="text-[10px] font-black text-slate-500 uppercase tracking-[0.4em]">No {activeTab === 'PENDING' ? 'Pending' : 'Resolved'} Requests</p>
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 pb-20">
            {currentRequests.map((req: any, i: number) => {
              const reqUniqueId = req.id || `request-${i}-${crypto.randomUUID()}`;
              return (
                <div key={reqUniqueId} className="bg-slate-900 border border-slate-800 p-6 rounded-[2rem] space-y-5 hover:border-indigo-500/30 transition-all shadow-xl group relative">
                <div className="flex justify-between items-start">
                  <div>
                    <p className="text-sm font-black text-white italic tracking-tight">{req.pic_name}</p>
                    <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">{req.schedule_date}</p>
                  </div>
                  <div className="flex items-center gap-2">
                    <div className={cn(
                      "px-3 py-1 text-[9px] font-black uppercase tracking-widest rounded-lg border",
                      req.status === 'Approved' ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/20" :
                      req.status === 'Rejected' ? "bg-rose-500/10 text-rose-500 border-rose-500/20" :
                      "bg-amber-500/10 text-amber-500 border-amber-500/20"
                    )}>
                      {req.status}
                    </div>
                    {activeTab === 'HISTORY' && (
                      <button 
                        onClick={() => setDeleteId(req.id)}
                        className="p-1.5 text-slate-600 hover:text-rose-500 hover:bg-rose-500/10 rounded-lg transition-all opacity-0 group-hover:opacity-100"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </div>
                </div>
                
                <div className="space-y-3">
                  <div className="flex items-center gap-4 bg-slate-950/50 p-4 rounded-2xl border border-slate-800/50">
                    <div className="flex-1 text-center">
                      <p className="text-[8px] text-slate-600 font-bold uppercase mb-1">From</p>
                      <span className="text-[10px] font-black text-slate-500">{req.original_status || 'KOSONG'}</span>
                    </div>
                    <ChevronRight className="w-4 h-4 text-slate-700" />
                    <div className="flex-1 text-center">
                      <p className="text-[8px] text-indigo-500 font-bold uppercase mb-1">To</p>
                      <span className="text-[10px] font-black text-indigo-400">{req.new_status}</span>
                    </div>
                  </div>

                  {req.swap_date && (
                    <div className="flex items-center gap-3 p-3 bg-indigo-500/5 rounded-xl border border-indigo-500/10">
                      <Layers className="w-3.5 h-3.5 text-indigo-400" />
                      <div>
                        <p className="text-[7px] text-slate-500 font-black uppercase tracking-widest">Two-Way Swap Detected</p>
                        <p className="text-[9px] text-slate-400 font-bold">Swap with {req.swap_date} ({req.swap_status})</p>
                      </div>
                    </div>
                  )}
                </div>

                <div className="p-4 bg-slate-950/30 rounded-2xl border border-white/5 h-20 overflow-y-auto">
                  <p className="text-[8px] text-slate-600 font-bold uppercase mb-1">Reason</p>
                  <p className="text-xs text-slate-400 italic font-medium leading-relaxed">"{req.reason}"</p>
                </div>

                {activeTab === 'HISTORY' && (
                  <div className="pt-4 border-t border-white/5 grid grid-cols-2 gap-4">
                    <div>
                      <p className="text-[7px] text-slate-600 font-bold uppercase tracking-widest">Processed By</p>
                      <p className="text-[9px] text-slate-400 font-mono mt-0.5">{req.processed_by || 'Admin'}</p>
                    </div>
                    <div className="text-right">
                      <p className="text-[7px] text-slate-600 font-bold uppercase tracking-widest">Action Date</p>
                      <p className="text-[9px] text-slate-400 font-mono mt-0.5">
                        {req.updated_at ? format(new Date(req.updated_at), 'dd/MM/yy HH:mm') : '-'}
                      </p>
                    </div>
                  </div>
                )}

                {req.status === 'Pending' && (
                  <div className="pt-2 flex gap-3 text-xs">
                    <button 
                      onClick={() => setConfirmData({ id: req.id, status: 'Rejected' })}
                      className="flex-1 py-3 bg-slate-800 hover:bg-rose-500/10 hover:text-rose-400 text-slate-500 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all"
                    >
                      Reject
                    </button>
                    <button 
                      onClick={() => setConfirmData({ id: req.id, status: 'Approved' })}
                      className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-[10px] font-black uppercase tracking-widest transition-all shadow-lg shadow-emerald-500/20"
                    >
                      Approve
                    </button>
                  </div>
                )}
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

function AuditLogView({ logs, projects, users }: { logs: AuditLog[], projects: Project[], users: AppUser[] }) {
  const [searchTerm, setSearchTerm] = useState('');
  const [dateFilter, setDateFilter] = useState('');
  const [selectedLog, setSelectedLog] = useState<AuditLog | null>(null);

  const filteredLogs = logs.filter(log => {
    const actorName = users.find(u => u.email === log.actor)?.name || log.actor || 'System';
    const matchesActor = actorName.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesDate = !dateFilter || (log.created_at && log.created_at.startsWith(dateFilter));
    return matchesActor && matchesDate;
  });

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between gap-4 bg-slate-900/50 p-4 rounded-2xl border border-slate-800">
        <div className="flex-1 relative">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
          <input 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="Search by Personnel Name..."
            className="w-full bg-slate-950/50 border border-slate-800 rounded-xl pl-10 pr-4 py-2.5 text-xs text-slate-300 outline-none focus:border-indigo-500 transition-all font-bold"
          />
        </div>
        <div className="flex items-center gap-3">
          <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Filter Date</label>
          <input 
            type="date"
            value={dateFilter}
            onChange={(e) => setDateFilter(e.target.value)}
            className="bg-slate-900 border border-slate-700 text-white rounded-xl px-4 py-2.5 text-xs outline-none focus:border-indigo-500 transition-all font-bold"
          />
        </div>
      </div>

      <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-2xl flex flex-col max-h-[calc(100vh-300px)]">
        <div className="flex-1 overflow-y-auto scrollbar-hide">
          <table className="w-full text-left">
            <thead className="sticky top-0 z-10 bg-slate-900 shadow-md">
              <tr className="border-b border-slate-800">
                <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Timestamp</th>
                <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Actor (Name)</th>
                <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Project Mapping</th>
                <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Operational Action</th>
                <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em] text-right">Insight</th>
              </tr>
            </thead>
        <tbody className="divide-y divide-slate-800/50 text-[12px] text-slate-400">
          {filteredLogs.map((log, i) => {
            const project = projects.find(p => p.id === log.project_id);
            const actorUser = users.find(u => u.email === log.actor);
            const logUniqueId = log.id || `audit-log-${i}-${crypto.randomUUID()}`;
            return (
              <tr key={`audit-table-row-${logUniqueId}`} className="hover:bg-white/[0.01] transition-colors group">
                    <td className="px-6 py-4 font-mono text-[10px] text-slate-500 whitespace-nowrap italic">
                      {log.created_at ? format(new Date(log.created_at), 'dd/MM/yyyy HH:mm') : 'N/A'}
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-2">
                        <div className="w-6 h-6 rounded-lg bg-indigo-500/10 border border-indigo-500/30 flex items-center justify-center font-black text-[9px] text-indigo-400">
                          {actorUser?.name?.charAt(0) || 'S'}
                        </div>
                        <div className="flex flex-col">
                          <span className="font-black text-slate-200 tracking-tight leading-none mb-1">{actorUser?.name || log.actor || 'System'}</span>
                          <span className="text-[9px] text-slate-600 font-bold uppercase tracking-widest">{log.actor}</span>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <span className="px-2 py-1 bg-slate-950 border border-slate-800 rounded text-[10px] text-indigo-400 uppercase font-black tracking-tighter">
                        {project?.name || 'Central Engine'}
                      </span>
                    </td>
                    <td className="px-6 py-4">
                       <span className={cn(
                         "text-[10px] font-black uppercase tracking-widest px-2 py-0.5 rounded",
                         log.action?.includes('Create') ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/20" :
                         log.action?.includes('Delete') ? "bg-rose-500/10 text-rose-400 border border-rose-500/20" :
                         "bg-indigo-500/10 text-indigo-400 border border-indigo-500/20"
                       )}>
                        {log.action || 'Unknown'}
                       </span>
                    </td>
                    <td className="px-6 py-4 text-right">
                       <button 
                         onClick={() => setSelectedLog(log)}
                         className="text-[10px] font-black uppercase text-white transition-colors tracking-widest px-4 py-1.5 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 rounded-lg shadow-lg shadow-indigo-500/20 transition-all active:scale-95"
                       >
                         View Details
                       </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      <AnimatePresence>
        {selectedLog && (
          <div className="fixed inset-0 z-[120] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setSelectedLog(null)}
              className="absolute inset-0 bg-slate-950/90 backdrop-blur-md"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="relative w-full max-w-4xl bg-slate-900 border border-slate-800 rounded-3xl shadow-2xl overflow-hidden max-h-[80vh] flex flex-col"
            >
              <div className="p-6 border-b border-slate-800 flex items-center justify-between">
                <div>
                  <h3 className="text-xl font-black text-white italic uppercase tracking-tighter">Payload Analysis</h3>
                  <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mt-1">Transaction ID: {selectedLog.id}</p>
                </div>
                <button onClick={() => setSelectedLog(null)} className="p-2 hover:bg-slate-800 rounded-xl transition-colors">
                  <X className="w-5 h-5 text-slate-500" />
                </button>
              </div>
              <div className="p-8 overflow-y-auto bg-slate-950/50 grow">
                <div className="space-y-6">
                  <label className="text-[10px] font-black text-white uppercase tracking-[0.2em] px-2 block border-l-2 border-indigo-500">Field Comparison Logic</label>
                  <div className="grid grid-cols-2 gap-4">
                    <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 text-[10px] font-bold text-slate-500 uppercase">Property</div>
                    <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 text-[10px] font-bold text-slate-500 uppercase">Transition</div>
                  </div>
                  {Object.keys({ ...(selectedLog.old_payload || {}), ...(selectedLog.new_payload || {}) }).map((key, ki) => {
                    // Exclude metadata fields
                    if (['updated_at', 'id', 'created_at', 'project_id', 'task_id', 'user_id', 'created_by_name'].includes(key)) return null;

                    const oldVal = (selectedLog.old_payload as any)?.[key];
                    const newVal = (selectedLog.new_payload as any)?.[key];
                    if (JSON.stringify(oldVal) === JSON.stringify(newVal)) return null;

                    const formatVal = (v: any) => {
                      if (v === null || v === undefined) return 'None';
                      if (typeof v === 'string') return v;
                      return JSON.stringify(v);
                    };

                    return (
                      <div key={`log-detail-payload-${key}-${ki}`} className="grid grid-cols-[150px_1fr] gap-4 items-center py-2 border-b border-white/5">
                        <div className="text-[10px] font-black text-slate-500 uppercase tracking-wider">{key.replace(/_/g, ' ')}</div>
                        <div className="flex items-center gap-3 overflow-hidden">
                           <div className="line-through text-rose-500/60 text-[10px] font-medium truncate">{formatVal(oldVal)}</div>
                           <div className="text-slate-700 font-bold">➔</div>
                           <div className="text-emerald-400 text-[11px] font-black italic">{formatVal(newVal)}</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

const DeferredTextarea = ({ value, onSave, className, placeholder }: { value: string, onSave: (v: string) => void, className?: string, placeholder?: string }) => {
  const [local, setLocal] = useState(value);
  useEffect(() => setLocal(value), [value]);

  return (
    <textarea 
      value={local}
      onChange={(e) => setLocal(e.target.value)}
      onBlur={() => { if (local !== value) onSave(local); }}
      className={className}
      placeholder={placeholder}
    />
  );
};

function AuditLogTable({ logs }: { logs: ProjectRescheduleLog[] }) {
  if (logs.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center py-20 text-slate-500 gap-3 border-2 border-dashed border-slate-800 rounded-3xl">
        <History className="w-8 h-8 opacity-20" />
        <p className="text-[10px] font-black uppercase tracking-widest">No Reschedule Events Logged</p>
      </div>
    );
  }

  return (
    <div className="overflow-hidden border border-white/5 rounded-2xl bg-slate-900/40">
      <div className="overflow-x-auto">
        <table className="w-full text-left border-collapse min-w-[800px]">
          <thead>
            <tr className="bg-slate-950/50 border-b border-white/5">
              <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-widest">Timestamp</th>
              <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-widest">Authorized By</th>
              <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">Previous Timeline</th>
              <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">Target Timeline</th>
              <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-widest">Reason / Justification</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-white/5">
            {logs.map((log, i) => (
              <tr key={log.id || `resched-log-${i}`} className="hover:bg-white/[0.02] transition-colors">
                <td className="px-6 py-4">
                  <div className="flex flex-col">
                    <span className="text-xs text-slate-200 font-mono italic">{format(new Date(log.created_at), 'MMM dd, yyyy')}</span>
                    <span className="text-[10px] text-slate-600 font-bold">{format(new Date(log.created_at), 'HH:mm:ss')}</span>
                  </div>
                </td>
                <td className="px-6 py-4">
                  <div className="flex items-center gap-2">
                    <div className="w-6 h-6 rounded-full bg-indigo-500/20 border border-indigo-500/30 flex items-center justify-center text-[10px] font-black text-indigo-400">
                      {log.changed_by?.charAt(0) || 'A'}
                    </div>
                    <span className="text-xs font-bold text-slate-300">{log.changed_by}</span>
                  </div>
                </td>
                <td className="px-6 py-4">
                   <div className="flex flex-col items-center gap-1 opacity-50">
                      <span className="text-[9px] text-slate-400 font-mono">{log.old_start_date}</span>
                      <ArrowDown className="w-3 h-3 text-slate-600" />
                      <span className="text-[9px] text-slate-400 font-mono">{log.old_end_date}</span>
                   </div>
                </td>
                <td className="px-6 py-4">
                   <div className="flex flex-col items-center gap-1">
                      <span className="text-[9px] text-emerald-400 font-black font-mono">{log.new_start_date}</span>
                      <ArrowDown className="w-3 h-3 text-indigo-500" />
                      <span className="text-[9px] text-emerald-400 font-black font-mono">{log.new_end_date}</span>
                   </div>
                </td>
                <td className="px-6 py-4">
                  <div className="p-3 bg-slate-950/50 rounded-xl border border-white/5 max-w-md">
                     <p className="text-xs text-slate-400 leading-relaxed italic">"{log.reason}"</p>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function GanttTree({ user, users, roots, map, tasks, projects, expandedRows, onToggleExpand, onUpdateTask, onOpenAudit, onAddSubTask, onDeleteTask }: any) {
  
  const renderTaskRows = (task: any, level: number = 0, index: number = 0) => {
    if (!task) return null;
    const isExpanded = expandedRows.has(task.id);
    const children = task.children || [];
    const isProject = !!task.isProject;
    const health = getTaskHealth(task);

    return (
      <React.Fragment key={`task-node-${task.id || `${level}-${index}`}`}>
        <tr 
          onClick={() => !isProject && children.length > 0 && onToggleExpand(task.id)}
          className={cn(
            "border-b border-white/5 transition-all group cursor-pointer",
            level === 0 ? "bg-slate-900/40" : "bg-slate-800/10",
            level === 1 && isExpanded ? "bg-slate-800/50" : "hover:bg-white/[0.02]",
            health === 'OVERDUE' && "border-l-4 border-l-rose-500 bg-rose-500/5",
            health === 'OVER SLA' && "border-l-4 border-l-amber-500 bg-amber-500/5"
          )}
        >
          {/* Node Selector / Title */}
          <td className="px-6 py-4">
            <div className={cn(
              "flex items-center gap-3",
              level === 1 ? "pl-10" : level > 1 ? "pl-16" : "pl-0"
            )}>
              {!isProject && children.length > 0 && (
                <span className="text-indigo-500 font-mono w-4">
                  {isExpanded ? "▼" : "▶"}
                </span>
              )}
              {isProject ? <FolderKanban className="w-4 h-4 text-indigo-400" /> : level > 0 ? <span className="text-indigo-500/60 font-black text-lg select-none leading-none">↳</span> : <Layers className="w-4 h-4 text-slate-500" />}
              <div className="flex-1 min-w-0" onClick={e => e.stopPropagation()}>
                <div className="flex items-center gap-2">
                  <EditableInput 
                    value={task.title} 
                    onSave={(v) => onUpdateTask(task.id, 'title', v)}
                    className="text-xs font-black text-white italic truncate tracking-tight uppercase bg-transparent outline-none border-none focus:text-indigo-400"
                  />
                  <HealthBadge health={health} />
                  {(() => {
                    const collisions = getCollision(task, tasks, projects);
                    return collisions ? <CollisionWarning key={`collision-${task.id}`} collisions={collisions} /> : null;
                  })()}
                </div>
                {level === 0 && !isProject && (() => {
                    const children = map.get(task.id) || [];
                    if (children.length > 0) {
                      const capacity = task.duration_hours || 0;
                      const used = children.reduce((sum: number, c: any) => sum + (c.duration_hours || 0), 0);
                      if (used > capacity) return <span className="text-[7px] font-black text-rose-500 uppercase mt-0.5">⚠️ Over-Allocated</span>;
                      if (used < capacity) return <span className="text-[7px] font-black text-amber-500 uppercase mt-0.5">⚠️ Under-Allocated</span>;
                      return <span className="text-[7px] font-black text-emerald-500 uppercase mt-0.5">✅ Alokasi Pas</span>;
                    }
                    return null;
                })()}
              </div>
            </div>
          </td>

          {/* Man-Hours */}
          <td className="px-2 py-4 text-center">
             <div className="flex flex-col items-center gap-1" onClick={e => e.stopPropagation()}>
                <div className="flex items-center gap-1">
                  <EditableInput 
                    type="number"
                    value={(task.duration_hours || 0).toString()} 
                    onSave={(v) => onUpdateTask(task.id, 'duration_hours', parseFloat(v) || 0)}
                    className="w-12 bg-slate-800 text-[10px] text-center rounded border border-slate-700 text-indigo-400 font-black"
                  />
                  <span className="text-[8px] text-slate-600 font-bold">h</span>
                </div>
                <div className="text-[7px] text-slate-500 font-black uppercase tracking-tighter bg-slate-900/50 px-1.5 py-0.5 rounded border border-white/5">
                  {formatWorkday(task.duration_hours || 0)}
                </div>
             </div>
          </td>

          {/* Assignee */}
          <td className="px-4 py-4" onClick={e => e.stopPropagation()}>
            <EditableInput 
              value={task.assignee || ''} 
              onSave={(v) => onUpdateTask(task.id, 'assignee', v)}
              className="bg-slate-800/80 text-[10px] font-bold text-center p-1 rounded border border-slate-700 text-slate-300 outline-none focus:border-indigo-500 w-full"
              placeholder="PIC"
            />
          </td>

          {/* Dev */}
          <td className="px-2 py-4" onClick={e => e.stopPropagation()}>
            {level > 0 && (
              <EditableInput 
                value={task.developer_name || ''} 
                onSave={(v) => onUpdateTask(task.id, 'developer_name', v)}
                className="bg-slate-950/50 border border-slate-800 text-[10px] px-1.5 py-1 rounded text-indigo-400/80 text-center hover:border-slate-700 transition-all font-mono w-full"
                placeholder="Dev"
              />
            )}
          </td>

          {/* QA */}
          <td className="px-2 py-4" onClick={e => e.stopPropagation()}>
            {level > 0 && (
              <EditableInput 
                value={task.qa_name || ''} 
                onSave={(v) => onUpdateTask(task.id, 'qa_name', v)}
                className="bg-slate-950/50 border border-slate-800 text-[10px] px-1.5 py-1 rounded text-purple-400/80 text-center hover:border-slate-700 transition-all font-mono w-full"
                placeholder="QA"
              />
            )}
          </td>

          {/* Dates */}
          <td 
            className="px-1 py-4 text-center cursor-pointer hover:bg-slate-800/80 transition-all group/date relative" 
            onClick={(e) => { 
              e.stopPropagation(); 
              const input = e.currentTarget.querySelector('input');
              if (input) {
                try { input.showPicker(); } catch(err) { input.focus(); }
              }
            }}
          >
            <div className="flex flex-col items-center">
              <input 
                type="date" 
                value={task.start_time ? format(new Date(task.start_time), "yyyy-MM-dd") : ''}
                onChange={(e) => onUpdateTask(task.id, 'start_time', new Date(e.target.value).toISOString())}
                className="bg-slate-900 border border-slate-700 text-[10px] text-white font-mono focus:text-indigo-400 outline-none w-full text-center cursor-pointer font-bold rounded px-1 py-0.5"
              />
              <span className="text-[7px] text-indigo-500/50 font-black uppercase opacity-0 group-hover/date:opacity-100 transition-all absolute -top-1">Start</span>
            </div>
          </td>

          <td 
            className="px-1 py-4 text-center cursor-pointer hover:bg-slate-800/80 transition-all group/date relative" 
            onClick={(e) => { 
              e.stopPropagation(); 
              const input = e.currentTarget.querySelector('input');
              if (input) {
                try { input.showPicker(); } catch(err) { input.focus(); }
              }
            }}
          >
            <div className="flex flex-col items-center">
              <input 
                type="date" 
                value={task.end_time ? format(new Date(task.end_time), "yyyy-MM-dd") : ''}
                onChange={(e) => onUpdateTask(task.id, 'end_time', new Date(e.target.value).toISOString())}
                className="bg-slate-900 border border-slate-700 text-[10px] text-white font-mono focus:text-indigo-400 outline-none w-full text-center cursor-pointer font-bold rounded px-1 py-0.5"
              />
              <span className="text-[7px] text-indigo-500/50 font-black uppercase opacity-0 group-hover/date:opacity-100 transition-all absolute -top-1">End</span>
            </div>
          </td>

          {/* Status */}
          <td className="px-4 py-4" onClick={e => e.stopPropagation()}>
            <div className="flex justify-center">
              <TaskStatusSelector 
                status={task.status || TaskStatus.TODO} 
                onUpdate={(v) => onUpdateTask(task.id, 'status', v)} 
              />
            </div>
          </td>

          {/* Fachrul Feedback */}
          <td className="px-4 py-4" onClick={e => e.stopPropagation()}>
             <div className="flex items-center gap-1.5 px-2 py-1.5 bg-slate-900/50 rounded-lg border border-slate-800/50 group-hover:border-slate-700/50 transition-colors">
                <div className="flex flex-col gap-0.5 min-w-[50px]">
                  <ApprovalBadge value={task.approval_fachrul} label="Fachrul" onUpdate={(v) => onUpdateTask(task.id, 'approval_fachrul', v)} />
                </div>
                <DeferredTextarea 
                   value={task.suggestion_fachrul || ''}
                   onSave={(v) => onUpdateTask(task.id, 'suggestion_fachrul', v)}
                   className="bg-slate-950 border border-slate-800 rounded px-2 py-1 text-[9px] text-slate-200 focus:ring-1 focus:ring-indigo-500/30 outline-none w-full min-h-[32px] max-h-[32px] resize-none scrollbar-hide font-medium"
                   placeholder="..."
                />
             </div>
          </td>

          {/* Barra Feedback */}
          <td className="px-4 py-4" onClick={e => e.stopPropagation()}>
             <div className="flex items-center gap-1.5 px-2 py-1.5 bg-slate-900/50 rounded-lg border border-slate-800/50 group-hover:border-slate-700/50 transition-colors">
                <div className="flex flex-col gap-0.5 min-w-[50px]">
                  <ApprovalBadge value={task.approval_barra} label="Barra" onUpdate={(v) => onUpdateTask(task.id, 'approval_barra', v)} />
                </div>
                <DeferredTextarea 
                   value={task.suggestion_barra || ''}
                   onSave={(v) => onUpdateTask(task.id, 'suggestion_barra', v)}
                   className="bg-slate-950 border border-slate-800 rounded px-2 py-1 text-[9px] text-slate-200 focus:ring-1 focus:ring-indigo-500/30 outline-none w-full min-h-[32px] max-h-[32px] resize-none scrollbar-hide font-medium"
                   placeholder="..."
                />
             </div>
          </td>

          {/* Hidden Actions Column if needed, or just combine in list */}
          <td className="px-6 py-4 text-right">
            <div className="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity" onClick={e => e.stopPropagation()}>
              {!isProject && level === 0 && (
                <button 
                  onClick={() => onAddSubTask(task.id, task.start_time, task.end_time)}
                  className="bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white text-[8px] font-black uppercase px-2 py-1 rounded shadow-lg shadow-indigo-500/20 transition-all border border-white/10 active:scale-95"
                >
                  + Breakdown
                </button>
              )}
              <button 
                onClick={() => onOpenAudit(task)}
                className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-500 hover:text-indigo-400 rounded-lg transition-all border border-slate-700"
              >
                <History className="w-3.5 h-3.5" />
              </button>
              <button 
                onClick={() => {
                  if (confirm('Decommission this node permanently?')) {
                    onDeleteTask(task.id);
                  }
                }}
                className="p-1.5 bg-slate-800 hover:bg-rose-500/20 text-slate-600 hover:text-rose-500 rounded-lg transition-all border border-slate-700"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </button>
            </div>
          </td>
        </tr>
        
        {isExpanded && (children || []).map((sub: any, sidx: number) => {
          const subKey = sub.id || `subtask-${task.id}-${sidx}-${crypto.randomUUID()}`;
          return <React.Fragment key={subKey}>{renderTaskRows(sub, level + 1, sidx)}</React.Fragment>;
        })}
      </React.Fragment>
    );
  };

  return (
    <div className="overflow-x-auto w-full scrollbar-thin scrollbar-thumb-slate-800 scrollbar-track-transparent">
      <table className="w-full text-left border-collapse min-w-[1600px]">
        <thead className="sticky top-0 z-40 bg-slate-900 border-b border-white/5">
          <tr className="shadow-xl">
            <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-[0.2em] min-w-[300px]">Hierarchy & Governance</th>
            <th className="px-2 py-4 text-[9px] font-black text-indigo-500 uppercase tracking-widest text-center w-24">Man-Hours</th>
            <th className="px-4 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-40 text-center">PIC</th>
            <th className="px-2 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-24 text-center">Dev</th>
            <th className="px-2 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-24 text-center">QA</th>
            <th className="px-4 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-24 text-center">Start</th>
            <th className="px-4 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-24 text-center">End</th>
            <th className="px-4 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-32 text-center">Status</th>
            <th className="px-4 py-4 text-[9px] font-black text-indigo-500 uppercase tracking-widest w-[250px]"><div className="flex items-center justify-center gap-1"><UserIcon className="w-2.5 h-2.5"/> Fachrul Feedback</div></th>
            <th className="px-4 py-4 text-[9px] font-black text-indigo-500 uppercase tracking-widest w-[250px]"><div className="flex items-center justify-center gap-1"><UserIcon className="w-2.5 h-2.5"/> Barra Feedback</div></th>
            <th className="px-6 py-4 text-[9px] font-black text-slate-500 uppercase tracking-widest w-40 text-right">Comm Ops</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-white/5">
          {roots.map((task: any, idx: number) => {
            const taskKey = task.id || `gantt-root-${idx}-${crypto.randomUUID()}`;
            return <React.Fragment key={taskKey}>{renderTaskRows(task, 0, idx)}</React.Fragment>;
          })}
        </tbody>
      </table>
    </div>
  );
}

function GanttTimeline({ user, scale, tasks, hierarchicalTasks, expandedRows, setTasks, projects, isGlobalView, onSetFocus }: { user: any, scale: ViewScale, tasks: Task[], hierarchicalTasks: any, expandedRows: Set<string>, setTasks: React.Dispatch<React.SetStateAction<Task[]>>, projects: Project[], isGlobalView?: boolean, onSetFocus?: (id: string) => void }) {
  // Use current month as center
  const now = new Date();
  
  const intervals = useMemo(() => {
    switch(scale) {
      case 'HOUR': {
        const start = startOfDay(now);
        return eachHourOfInterval({ start, end: endOfDay(now) });
      }
      case 'DAY': {
        const start = startOfMonth(now);
        // Show 45 days to cover the templates comfortably
        return Array.from({ length: 45 }).map((_, i) => addDays(start, i));
      }
      case 'WEEK': {
        const start = startOfMonth(now);
        return eachWeekOfInterval({ start, end: addMonths(start, 3) });
      }
      case 'MONTH': {
        return Array.from({ length: 12 }).map((_, i) => addMonths(startOfMonth(now), i));
      }
    }
  }, [scale]);

  const CELL_WIDTH = scale === 'HOUR' ? 80 : 100;

  return (
    <div className="flex h-full min-w-full">
      {/* Left Column: Task/Project Labels */}
      <div className="w-[200px] border-r border-slate-800 bg-slate-950/80 sticky left-0 z-30 flex flex-col shrink-0">
        <div className="h-10 border-b border-white/5 flex items-center px-4 bg-slate-900/90 sticky top-0">
          <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">Entity / Scope</span>
        </div>
        <div className="flex-1">
          {(hierarchicalTasks.roots || []).map((root: any, i: number) => {
             const isExpanded = expandedRows.has(root.id);
             const rootUniqueId = root.id || `root-fallback-${i}-${crypto.randomUUID()}`;
             return (
               <div key={`label-root-${rootUniqueId}`} className="flex flex-col">
                  <div className={cn(
                    "flex items-center px-4 border-b border-white/[0.02] text-white/70 font-bold text-[10px] uppercase truncate tracking-tighter",
                    isGlobalView ? "h-[64px]" : "h-[56px]"
                  )}>
                    {root.title}
                  </div>
                  {!isGlobalView && isExpanded && (root.children || []).map((child: any, ci: number) => {
                    const childUniqueId = child.id || `child-fallback-${ci}-${crypto.randomUUID()}`;
                    return (
                      <div key={`label-child-${childUniqueId}`} className="h-[48px] flex items-center px-4 pl-8 border-b border-white/[0.01] text-slate-500 text-[9px] font-medium truncate italic">
                        ↳ {child.title}
                      </div>
                    );
                  })}
               </div>
             );
          })}
        </div>
      </div>

      <div className="flex-1 overflow-visible">
        {/* Time Header matches sticky height of Tree Header */}
        <div className="flex sticky top-0 bg-slate-900/90 backdrop-blur-md border-b border-white/5 z-20 h-10">
          {(intervals || []).map((dt, i) => {
            if (!dt) return null;
            return (
              <div 
                key={dt.toISOString() || `interval-${i}`} 
                style={{ width: CELL_WIDTH }}
                className="flex-shrink-0 border-r border-white/5 flex flex-col justify-center px-3"
              >
                <span className="text-[10px] font-bold text-slate-500 uppercase tracking-tighter">
                  {scale === 'HOUR' ? format(dt, 'HH:mm') : format(dt, 'MMM dd')}
                </span>
                <span className="text-[9px] text-slate-600 font-medium">
                   {scale === 'HOUR' ? 'TODAY' : format(dt, 'EEEE')}
                </span>
              </div>
            );
          })}
        </div>

        {/* Grid Lines */}
        <div className="relative">
          <div className="absolute inset-0 flex pointer-events-none">
            {(intervals || []).map((dt, i) => (
              <div key={dt?.toISOString() || `grid-${i}`} style={{ width: CELL_WIDTH }} className="flex-shrink-0 border-r border-white/5" />
            ))}
          </div>

          {/* Task Bars aligned with Tree rows */}
          <div className="relative z-10">
            {(hierarchicalTasks.roots || []).map((root: any, i: number) => {
               const isExpanded = expandedRows.has(root.id);
               const rootUniqueId = root.id || `timeline-root-fallback-${i}`;
               return (
                 <div key={`timeline-root-${rootUniqueId}-${i}`} className="flex flex-col">
                    {/* L1 or Project Bar */}
                    <div className={cn(
                      "flex items-center border-b border-white/[0.02]",
                      isGlobalView ? "h-[64px]" : "h-[56px]"
                    )}>
                      <GanttBar 
                        user={user}
                        task={root} 
                        tasks={tasks}
                        projects={projects}
                        setTasks={setTasks}
                        scale={scale} 
                        intervals={intervals} 
                        cellWidth={CELL_WIDTH} 
                        isLevel1={true}
                        isProjectBar={isGlobalView}
                        onSetFocus={onSetFocus}
                      />
                    </div>
                    
                    {!isGlobalView && isExpanded && (root.children || []).map((child: any, ci: number) => {
                      const childUniqueId = child.id || `timeline-child-fallback-${ci}`;
                      return (
                        <div key={`timeline-child-${childUniqueId}-${ci}`} className="h-[48px] flex items-center border-b border-white/[0.01]">
                          <GanttBar 
                            user={user}
                            task={child} 
                            tasks={tasks}
                            projects={projects}
                            setTasks={setTasks}
                            scale={scale} 
                            intervals={intervals} 
                            cellWidth={CELL_WIDTH} 
                            isLevel1={false}
                          />
                        </div>
                      );
                    })}
                 </div>
               );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}

function GanttBar({ user, task, tasks, projects, setTasks, scale, intervals, cellWidth, isLevel1, isProjectBar, onSetFocus }: any) {
  const [isDragging, setIsDragging] = useState(false);
  const [dragOffset, setDragOffset] = useState(0);
  const [showPopover, setShowPopover] = useState(false);

  const start = task?.start_time ? new Date(task.start_time) : new Date();
  
  const totalDurationMinutes = ((task.duration_hours || 0) * 60) + (task.duration_minutes || 0);
  const end = new Date(start.getTime() + totalDurationMinutes * 60000);
  const anchor = intervals && intervals[0] ? intervals[0] : start;

  const getOffset = () => {
    if (!anchor || !start) return 0;
    if (scale === 'HOUR') {
      const diffMinutes = (start.getTime() - anchor.getTime()) / 60000;
      return (diffMinutes / 60) * cellWidth;
    }
    const diff = differenceInDays(start, anchor);
    return Math.max(0, diff * cellWidth);
  };

  const getWidth = () => {
    if (scale === 'HOUR') {
      return Math.max(CELL_MIN_WIDTH, (totalDurationMinutes / 60) * cellWidth);
    }
    return Math.max(CELL_MIN_WIDTH, (totalDurationMinutes / 1440) * cellWidth);
  };

  const CELL_MIN_WIDTH = 10;

  const handlePointerDown = (e: React.PointerEvent) => {
    if (isProjectBar) return; // Project summaries are read-only
    if (task.status === TaskStatus.DONE) return;
    setIsDragging(true);
    setShowPopover(false);
    const startX = e.clientX;
    
    const handleMove = (moveEvent: PointerEvent) => {
      const delta = moveEvent.clientX - startX;
      const step = cellWidth / (scale === 'HOUR' ? 1 : 24);
      const snappedDelta = Math.round(delta / step) * step;
      setDragOffset(snappedDelta);
    };

    const handleUp = async () => {
      setIsDragging(false);
      setDragOffset(0);
      window.removeEventListener('pointermove', handleMove);
      window.removeEventListener('pointerup', handleUp);

      if (dragOffset !== 0) {
        const units = dragOffset / cellWidth;
        const hoursDiff = scale === 'HOUR' ? units : units * 24;
        const newStart = addHours(start, Math.round(hoursDiff));
        
        const updatedTask = { 
          ...task, 
          start_time: newStart.toISOString(),
          end_time: new Date(newStart.getTime() + totalDurationMinutes * 60000).toISOString() 
        };

        setTasks((prev: Task[]) => prev.map((t: any) => t.id === task.id ? updatedTask : t));
        
        try {
          await taskService.updateTask(task.id, {
            start_time: updatedTask.start_time,
            end_time: updatedTask.end_time
          }, user?.email || 'Administrator');
        } catch (err) {
          console.error('Update failed:', err);
        }
      }
    };

    window.addEventListener('pointermove', handleMove);
    window.addEventListener('pointerup', handleUp);
  };

  const isApproved = task.status === TaskStatus.DONE;
  const isNeedsRevision = task.status === 'Needs Revision';
  const collisions = getCollision(task, tasks, projects);
  const health = getTaskHealth(task);

  const timeRemaining = Math.max(0, (end.getTime() - new Date().getTime()) / 3600000);
  const timeProgress = Math.min(100, Math.max(0, ((new Date().getTime() - start.getTime()) / 60000 / totalDurationMinutes) * 100));

  return (
    <div className={cn("relative flex items-center", isProjectBar ? "h-16" : (isLevel1 ? "h-14" : "h-10"))}>
      <motion.div
        layoutId={task.id}
        initial={false}
        animate={{ 
          x: getOffset() + dragOffset, 
          width: getWidth(),
          opacity: 1,
          scale: isDragging ? 1.02 : 1,
          zIndex: isDragging ? 50 : 10
        }}
        onPointerDown={handlePointerDown}
        onClick={(e) => {
          if (isProjectBar && onSetFocus) {
            e.stopPropagation();
            onSetFocus(task.id);
          } else {
            setShowPopover(!showPopover);
          }
        }}
        className={cn(
          "absolute rounded-lg flex items-center px-4 shadow-lg group/bar select-none transition-all duration-300",
          isProjectBar 
            ? "h-10 outline outline-2 outline-indigo-500/20 cursor-pointer hover:opacity-80 hover:scale-[1.02] active:scale-95" 
            : cn(isLevel1 ? "h-8" : "h-5", "cursor-move"),
          isApproved 
            ? "bg-emerald-500/20 border border-emerald-500/30 text-emerald-300"
            : isNeedsRevision 
              ? "bg-amber-500/20 border border-amber-500/30 text-amber-300"
              : "bg-indigo-500/20 border border-indigo-500/30 text-indigo-100",
          isDragging && "shadow-indigo-500/20 ring-2 ring-indigo-500/50",
          collisions && "ring-2 ring-red-500 border-red-500/50 animate-pulse",
          health === 'OVERDUE' && "bg-rose-500/40 border-rose-500 shadow-[0_0_15px_rgba(244,63,94,0.4)] text-rose-100 font-bold",
          health === 'OVER SLA' && "bg-amber-500/40 border-amber-500 shadow-[0_0_15px_rgba(245,158,11,0.4)] text-amber-100 font-bold"
        )}
      >
        <div className={cn(
          "absolute left-0 w-1 h-full rounded-l-lg",
          health === 'OVERDUE' ? "bg-rose-500" : (health === 'OVER SLA' ? "bg-amber-500" : (isApproved ? "bg-emerald-500" : "bg-indigo-500"))
        )} />
        
        <span className={cn(
          "font-black whitespace-nowrap overflow-hidden transition-all group-hover/bar:scale-105 uppercase tracking-tighter",
          isLevel1 || isProjectBar ? "text-[10px]" : "text-[8px]"
        )}>
           {task.title}
        </span>

        <AnimatePresence>
          {showPopover && (
            <motion.div 
              initial={{ opacity: 0, y: 10, scale: 0.9 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 10, scale: 0.9 }}
              className="absolute bottom-full mb-3 left-1/2 -translate-x-1/2 w-64 bg-slate-900 border border-slate-700 rounded-xl p-4 shadow-2xl z-[100] backdrop-blur-xl"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="space-y-4">
                <div className="flex justify-between items-start">
                  <h4 className="text-[11px] font-black text-indigo-400 uppercase tracking-[0.1em]">{task.title}</h4>
                  <HealthBadge health={health} />
                </div>
                
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <p className="text-[8px] font-bold text-slate-500 uppercase tracking-widest">Start Time</p>
                    <p className="text-[11px] font-mono text-slate-200">{format(start, 'MMM dd, HH:mm')}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[8px] font-bold text-slate-500 uppercase tracking-widest">Target End</p>
                    <p className="text-[11px] font-mono text-slate-200">{format(end, 'MMM dd, HH:mm')}</p>
                  </div>
                </div>

                <div className="space-y-2">
                  <div className="flex justify-between text-[8px] font-bold uppercase tracking-widest">
                    <span className="text-slate-500">Execution Progress</span>
                    <span className="text-indigo-400">{Math.round(timeProgress)}%</span>
                  </div>
                  <div className="h-1.5 bg-slate-800 rounded-full overflow-hidden">
                    <div className="h-full bg-indigo-500 shadow-[0_0_8px_#6366f1]" style={{ width: `${timeProgress}%` }} />
                  </div>
                </div>

                <div className="pt-2 border-t border-slate-800 flex justify-between items-center">
                  <span className="text-[9px] font-bold text-slate-500 uppercase">Remaining</span>
                  <span className={cn("text-[11px] font-black", timeRemaining < 12 ? "text-rose-400" : "text-emerald-400")}>
                    {Math.round(timeRemaining)} HOURS
                  </span>
                </div>
              </div>
              <div className="absolute top-full left-1/2 -translate-x-1/2 border-8 border-transparent border-t-slate-900" />
            </motion.div>
          )}
        </AnimatePresence>
      </motion.div>
    </div>
  );
}
