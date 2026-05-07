export enum TaskStatus {
  TODO = 'Todo',
  ON_HOLD = 'On Hold',
  ON_PROGRESS = 'On Progress',
  DONE = 'Done'
}

export type Status = TaskStatus; // For backward compatibility as literal alias

export enum ProjectStatus {
  ACTIVE = 'Active',
  DONE = 'Done',
  PROJECT_LATE = 'Project Late',
  FSD_PROGRESS = 'FSD on Progress',
  FSD_REVIEW = 'FSD on Review',
  SIT_PROGRESS = 'SIT on Progress',
  ON_HOLD = 'On Hold',
  UAT_PROGRESS = 'UAT on Progress'
}

export interface Project {
  id: string;
  name: string;
  status: ProjectStatus;
  leader_email?: string;
  pic_name?: string; // Free text for Project Lead/PIC
  start_date?: string;
  end_date?: string;
  created_at: string;
  updated_at: string;
}

export interface AppUser {
  id: string;
  name: string;
  email: string;
  password?: string;
  access_level: string; // Dropdown: Superadmin/Admin/PIC/etc.
  role: string;         // Input Text: Developer/QA/etc.
  created_at: string;
  updated_at?: string;
}

export interface Task {
  id: string;
  parent_id: string | null;
  project_id: string | null;
  title: string;
  assignee: string; // Used as display name (free text)
  developer_name?: string;
  qa_name?: string;
  start_time: string; // ISO string
  end_time: string;   // ISO string
  start_hour: number;  // 0-23
  start_minute: number; // 0-59
  duration_hours: number;
  duration_minutes: number;
  target_sla_date: string | null;
  created_by_name?: string;
  status: TaskStatus;
  approval_fachrul: string | null;
  suggestion_fachrul: string | null;
  approval_barra: string | null;
  suggestion_barra: string | null;
  created_at: string;
  updated_at: string;
}

export interface AuditLog {
  id: string;
  task_id?: string;
  project_id?: string;
  user_id?: string;
  actor: string;
  action: string;
  old_payload: any;
  new_payload: any;
  created_at: string;
  updated_at?: string;
}

export type ViewScale = 'MONTH' | 'WEEK' | 'DAY' | 'HOUR';
export type AppView = 'PROJECTS' | 'KANBAN' | 'PERSONEL' | 'AUDIT' | 'GANTT_DETAIL' | 'SCHEDULE' | 'RESCHEDULE';

export interface Schedule {
  id: string;
  pic_name: string;
  schedule_date: string; // ISO date string
  status: string;
  created_at: string;
  updated_at: string;
}

export interface RescheduleRequest {
  id: string;
  pic_name: string;
  schedule_date: string;
  original_status: string;
  new_status: string;
  reason: string;
  requested_by: string; // email or name
  status: 'Pending' | 'Approved' | 'Rejected';
  swap_date?: string;
  swap_status?: string;
  created_at: string;
  updated_at: string;
}

export interface ProjectRescheduleLog {
  id: string;
  project_id: string;
  changed_by: string;
  old_start_date: string;
  old_end_date: string;
  new_start_date: string;
  new_end_date: string;
  reason: string;
  created_at: string;
}
