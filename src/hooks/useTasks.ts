import { useState, useEffect, useCallback } from 'react';
import { Task, AuditLog } from '../types';
import { taskService } from '../services/taskService';
import { supabase } from '../lib/supabase';

export function useTasks() {
  const [tasks, setTasks] = useState<Task[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const fetchTasks = useCallback(async () => {
    try {
      setLoading(true);
      const data = await taskService.getTasks();
      setTasks(data);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchTasks();

    // 1. Initialize the channel and chain ALL events BEFORE subscribing
    const taskChannel = supabase
      .channel('public:tasks_hook')
      .on(
        'postgres_changes', 
        { event: '*', schema: 'public', table: 'tasks' }, 
        (payload) => {
          console.log('PAK TARNO (Hook): Realtime update:', payload.eventType);
          fetchTasks();
        }
      )
      .subscribe((status) => {
        if (status === 'SUBSCRIBED') {
          console.log('PAK TARNO (Hook): Successfully connected to Realtime');
        }
      });

    // 2. Bulletproof Cleanup
    return () => {
      if (taskChannel) {
        supabase.removeChannel(taskChannel);
      }
    };
  }, [fetchTasks]);

  return { tasks, loading, error, refresh: fetchTasks };
}
