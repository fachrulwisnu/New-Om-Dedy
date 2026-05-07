import { supabase } from '../lib/supabase';
import { Task, AuditLog, TaskStatus, ProjectStatus, Project, AppUser, Schedule, RescheduleRequest, ProjectRescheduleLog } from '../types';
import { format, endOfMonth } from 'date-fns';

/**
 * Service to manage tasks and their audit trails.
 * Enforces the requirement that every mutation triggers an audit log.
 */
export const taskService = {
  // --- Tasks ---
  async getTasks(projectId?: string): Promise<Task[]> {
    let query = supabase.from('tasks').select('*').order('created_at', { ascending: true });
    if (projectId) {
      query = query.eq('project_id', projectId);
    }
    const { data, error } = await query;
    if (error) throw error;
    return data || [];
  },

  async getAllTasks(): Promise<Task[]> {
    const { data, error } = await supabase
      .from('tasks')
      .select('*');
    if (error) throw error;
    return data || [];
  },

  // --- Projects ---
  async getProjects(): Promise<Project[]> {
    const { data, error } = await supabase
      .from('projects')
      .select('*')
      .order('created_at', { ascending: false });
    if (error) throw error;
    return data || [];
  },

  async createProject(project: Partial<Project>, actor: string): Promise<Project> {
    const { data: { user } } = await supabase.auth.getUser();
    const finalActor = user?.email || actor;

    const { data, error } = await supabase
      .from('projects')
      .insert([{
        ...project,
        leader_email: project.leader_email || finalActor,
        pic_name: project.pic_name || null
      }])
      .select()
      .single();
    if (error) throw error;
    await this.logAudit({ project_id: data.id, actor: finalActor, action: 'Created Project', newValue: data });
    return data;
  },

  async deleteProject(id: string, actor: string): Promise<void> {
    const { error } = await supabase.from('projects').delete().eq('id', id);
    if (error) throw error;
    await this.logAudit({ project_id: id, actor, action: 'Deleted Project' });
  },

  // --- Users ---
  async getUsers(): Promise<AppUser[]> {
    const { data, error } = await supabase
      .from('users')
      .select('*')
      .order('created_at', { ascending: false });
    if (error) throw error;
    return data || [];
  },

  async createUser(user: Partial<AppUser>, actor: string): Promise<AppUser> {
    const payload = {
      ...user,
      access_level: user.access_level || 'PIC',
      role: user.role || 'Staff',
      updated_at: new Date().toISOString()
    };

    const { data, error } = await supabase
      .from('users')
      .insert([payload])
      .select()
      .single();
    
    if (error) {
      if (error.code === '23505') {
        const err = new Error('Email ini sudah terdaftar dalam sistem.');
        (err as any).code = '23505';
        throw err;
      }
      throw error;
    }
    await this.logAudit({ user_id: data.id, actor, action: 'Created User', newValue: data });
    return data;
  },

  async deleteUser(id: string, actor: string): Promise<void> {
    const { error } = await supabase.from('users').delete().eq('id', id);
    if (error) throw error;
    await this.logAudit({ user_id: id, actor, action: 'Deleted User' });
  },

  async updateUser(id: string, updates: Partial<AppUser>, actor: string): Promise<AppUser> {
    const { data: existing, error: fetchError } = await supabase
      .from('users')
      .select('*')
      .eq('id', id)
      .single();
    if (fetchError) throw fetchError;

    const { data: updated, error: updateError } = await supabase
      .from('users')
      .update({ ...updates, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select()
      .single();
    if (updateError) throw updateError;

    await this.logAudit({ user_id: id, actor, action: 'Updated User', oldValue: existing, newValue: updated });
    return updated;
  },

  // --- Reschedule Requests ---
  async getRescheduleRequests(): Promise<any[]> {
    const { data, error } = await supabase
      .from('reschedule_requests')
      .select('*')
      .order('created_at', { ascending: false });
    if (error) {
      console.error('Error fetching reschedule requests:', error);
      throw error;
    }
    return data || [];
  },

  async checkExistingRescheduleRequest(pic_name: string, schedule_date: string): Promise<boolean> {
    const { data, error } = await supabase
      .from('reschedule_requests')
      .select('id')
      .eq('pic_name', pic_name)
      .eq('schedule_date', schedule_date)
      .eq('status', 'Pending')
      .limit(1);
    
    if (error) {
      console.error('Error checking existing reschedule request:', error);
      return false; 
    }
    return (data && data.length > 0);
  },

  async createRescheduleRequest(payload: any | any[]): Promise<void> {
    const payloads = Array.isArray(payload) ? payload : [payload];
    const formattedPayloads = payloads.map(p => ({
      pic_name: p.pic_name,
      schedule_date: p.schedule_date,
      original_status: p.original_status,
      new_status: p.new_status,
      reason: p.reason,
      requested_by: p.requested_by,
      status: p.status || 'Pending',
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    }));

    const { error } = await supabase
      .from('reschedule_requests')
      .insert(formattedPayloads);
    
    if (error) {
      console.error('Error creating reschedule request:', error);
      throw error;
    }
  },

  async deleteRescheduleRequest(id: string, actor: string): Promise<void> {
    const { error } = await supabase
      .from('reschedule_requests')
      .delete()
      .eq('id', id);
    if (error) throw error;
    await this.logAudit({ actor, action: 'Deleted Reschedule Request', newValue: { id } });
  },

  async updateRescheduleRequestStatus(id: string, status: 'Approved' | 'Rejected', actor: string): Promise<void> {
    const { data: request, error: fetchError } = await supabase
      .from('reschedule_requests')
      .select('*')
      .eq('id', id)
      .single();
    
    if (fetchError) throw fetchError;

    const { error } = await supabase
      .from('reschedule_requests')
      .update({ status, updated_at: new Date().toISOString() })
      .eq('id', id);
    
    if (error) throw error;

    if (status === 'Approved') {
      // If approved, update the actual schedule
      const upserts = [{
        pic_name: request.pic_name,
        schedule_date: request.schedule_date,
        status: request.new_status
      }];

      // Handle Two-Way Swap
      if (request.swap_date && request.swap_status) {
        upserts.push({
          pic_name: request.pic_name,
          schedule_date: request.swap_date,
          status: request.swap_status
        });
      }

      await this.upsertSchedules(upserts);
    }

    await this.logAudit({ 
      actor, 
      action: `Reschedule Request ${status}`, 
      oldValue: request,
      newValue: { ...request, status }
    });
  },

  async getProjectRescheduleLogs(projectId: string): Promise<ProjectRescheduleLog[]> {
    const { data, error } = await supabase
      .from('project_reschedule_logs')
      .select('*')
      .eq('project_id', projectId)
      .order('created_at', { ascending: false });
    if (error) throw error;
    return data || [];
  },

  async createProjectRescheduleLog(log: Partial<ProjectRescheduleLog>): Promise<void> {
    const sanitizeDate = (d: any) => (d === 'N/A' || !d || d === 'null' || d === 'undefined') ? null : d;
    const sanitizedLog = {
      ...log,
      old_start_date: sanitizeDate(log.old_start_date),
      old_end_date: sanitizeDate(log.old_end_date),
      new_start_date: sanitizeDate(log.new_start_date),
      new_end_date: sanitizeDate(log.new_end_date),
    };
    const { error } = await supabase
      .from('project_reschedule_logs')
      .insert([sanitizedLog]);
    if (error) throw error;
    
    await this.logAudit({ 
      project_id: log.project_id, 
      actor: log.changed_by || 'System', 
      action: 'Project Rescheduled', 
      newValue: sanitizedLog 
    });
  },

  async updateProject(id: string, updates: Partial<Project>, actor: string): Promise<Project> {
    const { data: existing, error: fetchError } = await supabase
      .from('projects')
      .select('*')
      .eq('id', id)
      .single();
    if (fetchError) throw fetchError;

    const { data: updated, error: updateError } = await supabase
      .from('projects')
      .update({ ...updates, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select()
      .single();
    if (updateError) throw updateError;

    await this.logAudit({ project_id: id, actor, action: 'Updated Project', oldValue: existing, newValue: updated });
    return updated;
  },

  // --- Audit Logs ---
  async getAuditLogs(params?: { taskId?: string, projectId?: string, userId?: string }): Promise<AuditLog[]> {
    let query = supabase.from('audit_logs').select('*').order('created_at', { ascending: false });
    if (params?.taskId) query = query.eq('task_id', params.taskId);
    if (params?.projectId) query = query.eq('project_id', params.projectId);
    if (params?.userId) query = query.eq('user_id', params.userId);
    
    const { data, error } = await query;
    if (error) throw error;
    return data || [];
  },

  async logAudit({ task_id, project_id, user_id, actor, action, oldValue, newValue }: { 
    task_id?: string, 
    project_id?: string, 
    user_id?: string, 
    actor: string, 
    action: string, 
    oldValue?: any, 
    newValue?: any 
  }) {
    const { error } = await supabase
      .from('audit_logs')
      .insert([{
        task_id,
        project_id,
        user_id,
        actor,
        action,
        action_type: action, // Fix Error 23502 by explicitly mapping action to action_type
        old_payload: oldValue || null,
        new_payload: newValue || null
      }]);
    
    if (error) console.error('Failed to log audit:', JSON.stringify(error));
  },

  async deleteTask(id: string, actor: string) {
    const { data: existing } = await supabase.from('tasks').select('*').eq('id', id).single();
    
    // Level 1 Delete: Cascade to all children (Level 2)
    const { data: children } = await supabase.from('tasks').select('id').eq('parent_id', id);
    if (children && children.length > 0) {
      const childIds = children.map(c => c.id);
      await supabase.from('tasks').delete().in('id', childIds);
    }

    const { error } = await supabase.from('tasks').delete().eq('id', id);
    if (error) throw error;

    await this.logAudit({ 
      task_id: id, 
      project_id: existing?.project_id,
      actor, 
      action: 'DELETED', 
      oldValue: existing 
    });
  },

  async createTask(task: Partial<Task>, actor: string): Promise<Task> {
    const { data: { user } } = await supabase.auth.getUser();
    const created_by = user?.email || actor;
    
    // Safety calculation for end_time if start_time and durations are present
    let calculatedEndTime = task.end_time;
    if (!calculatedEndTime && task.start_time) {
      try {
        const start = new Date(task.start_time);
        const mins = ((task.duration_hours || 0) * 60) + (task.duration_minutes || 0);
        if (mins > 0) {
          calculatedEndTime = new Date(start.getTime() + mins * 60000).toISOString();
        }
      } catch (e) {
        console.warn("Could not calculate end_time", e);
      }
    }

    const payload: any = {
      ...task,
      assignee: task.assignee || created_by,
      created_by_name: created_by,
      start_hour: parseInt(String(task.start_hour ?? 0)) || 0,
      start_minute: parseInt(String(task.start_minute ?? 0)) || 0,
      duration_hours: parseInt(String(task.duration_hours ?? 0)) || 0,
      duration_minutes: parseInt(String(task.duration_minutes ?? 0)) || 0,
      status: TaskStatus.ON_PROGRESS
    };

    if (calculatedEndTime) {
      payload.end_time = calculatedEndTime;
    }

    const { data, error } = await supabase
      .from('tasks')
      .insert([payload])
      .select()
      .single();

    if (error) throw error;

    await this.logAudit({ 
      task_id: data.id, 
      project_id: data.project_id || undefined, 
      actor, 
      action: 'Created Task', 
      newValue: data 
    });
    return data;
  },

  async updateTask(id: string, updates: Partial<Task>, actor: string): Promise<Task> {
    const { data: existing, error: fetchError } = await supabase
      .from('tasks')
      .select('*')
      .eq('id', id)
      .single();

    if (fetchError) throw fetchError;

    const { data: updated, error: updateError } = await supabase
      .from('tasks')
      .update({ ...updates, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select()
      .single();

    if (updateError) throw updateError;

    await this.logAudit({ 
      task_id: id, 
      project_id: updated.project_id || undefined, 
      actor, 
      action: 'Updated Task', 
      oldValue: existing, 
      newValue: updated 
    });
    
    return updated;
  },

  // --- Om Dedy Schedules ---
  async getSchedules(monthStart?: Date): Promise<Schedule[]> {
    let query = supabase.from('schedules').select('*');
    
    if (monthStart) {
      const start = format(monthStart, 'yyyy-MM-01');
      const end = format(endOfMonth(monthStart), 'yyyy-MM-dd');
      query = query.gte('schedule_date', start).lte('schedule_date', end);
    }

    const { data, error } = await query.order('schedule_date', { ascending: true });
    if (error) throw error;
    return data || [];
  },

  async upsertSchedules(schedules: Partial<Schedule>[]): Promise<void> {
    const { error } = await supabase
      .from('schedules')
      .upsert(schedules, { onConflict: 'pic_name,schedule_date' });
    if (error) throw error;
  },

  async seedSampleData(actor: string): Promise<void> {
    const now = new Date();
    const day = (d: number) => new Date(now.getTime() + (d - 1) * 24 * 60 * 60 * 1000).toISOString();
    
    // 1. Create Projects
    const prj1 = await this.createProject({ name: 'Digitalisasi SPPK Phase 2', status: ProjectStatus.ACTIVE }, actor);
    const prj2 = await this.createProject({ name: 'Enhancement Asset Management', status: ProjectStatus.ACTIVE }, actor);

    // 2. Create Users
    await this.createUser({ name: 'Hari', access_level: 'PIC', role: 'PIC' }, actor);
    await this.createUser({ name: 'William', access_level: 'PIC', role: 'Developer' }, actor);
    await this.createUser({ name: 'Salvador', access_level: 'PIC', role: 'Developer' }, actor);
    await this.createUser({ name: 'Syahid', access_level: 'Admin', role: 'QA' }, actor);
    await this.createUser({ name: 'Danuh', access_level: 'PIC', role: 'PIC' }, actor);

    // 3. Create Tasks for Project 1
    const p1_root = await this.createTask({
      title: 'Project Timeline',
      assignee: 'Head Manager',
      start_time: day(1),
      end_time: day(34),
      project_id: prj1.id
    }, actor);

    await this.createTask({ title: 'Breakdown FSD', parent_id: p1_root.id, project_id: prj1.id, assignee: 'Hari', start_time: day(1), end_time: day(2) }, actor);
    await this.createTask({ title: 'Create FSD', parent_id: p1_root.id, project_id: prj1.id, assignee: 'Hari', start_time: day(3), end_time: day(5) }, actor);
    await this.createTask({ 
      title: 'Feature: Auth Module', 
      parent_id: p1_root.id, 
      project_id: prj1.id, 
      assignee: 'William', 
      developer_name: 'William',
      start_time: day(8), 
      end_time: day(15) 
    }, actor);

    // 4. Create Tasks for Project 2 (Collision)
    const p2_root = await this.createTask({
      title: 'Enhancement Timeline',
      assignee: 'Manager 2',
      start_time: day(10),
      end_time: day(20),
      project_id: prj2.id
    }, actor);

    await this.createTask({ 
      title: 'Asset Module Dev (Collision with William)', 
      parent_id: p2_root.id, 
      project_id: prj2.id, 
      assignee: 'William', 
      developer_name: 'William',
      start_time: day(12), 
      end_time: day(18) 
    }, actor);
  }
};
