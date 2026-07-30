'use client';

import React, { useState, useEffect } from 'react';
import { supabase } from '@/lib/supabase';
import * as XLSX from 'xlsx';
import { CheckCircle2, Circle, Clock, Upload, Plus, AlertCircle, Users, User, Flame, Edit3, Save, X, CheckCheck, FileText, Trash2, LayoutList, Tag, Calendar, Folder } from 'lucide-react';

interface EKOTask {
  id: string;
  title: string;
  subtask: string;
  assigned_to: string;
  team: string;
  scheduled_at: string | null; // Now allows null/blank dates
  is_completed: boolean;
  notes: string;
}

// Safe string helper
const safeLower = (val: any): string => String(val ?? '').toLowerCase().trim();

// Strict Category Header Detection
const detectCategoryHeader = (cellStrings: string[]): string | null => {
  const nonEmp = cellStrings.filter((s) => s !== '');
  if (nonEmp.length === 0 || nonEmp.length > 3) return null;

  for (const str of nonEmp) {
    const lower = str.toLowerCase().trim();
    
    let foundCat: string | null = null;
    if (lower === 'topics brainstorming' || lower === 'topics brainstorm' || lower === 'topic brainstorming' || lower === 'brainstorming') {
      foundCat = 'Topics Brainstorming';
    } else if (lower === 'qa topics content' || lower === 'qa topics' || lower === 'qa topic content' || lower === 'qa content') {
      foundCat = 'QA Topics Content';
    } else if (lower === 'hackathon' || lower === 'hackathon topics' || lower === 'hackathon activities') {
      foundCat = 'Hackathon';
    } else if (lower === 'india swag' || lower === 'swag' || lower === 'swag items' || lower === 'swag activities') {
      foundCat = 'India Swag';
    } else if (lower === 'miscellaneous' || lower === 'misc' || lower === 'miscellaneous tasks') {
      foundCat = 'Miscellaneous';
    }

    if (foundCat && str.length < 40) {
      return foundCat;
    }
  }

  return null;
};

// Safe Optional Date Converter
const parseExcelDate = (val: any): string | null => {
  if (!val || String(val).trim() === '') return null; // Keep it blank if empty
  
  try {
    if (val instanceof Date && !isNaN(val.getTime())) {
      return val.toISOString();
    }
    if (typeof val === 'number') {
      const parsed = XLSX.SSF.parse_date_code(val);
      if (parsed) {
        const d = new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, parsed.S || 0);
        if (!isNaN(d.getTime())) return d.toISOString();
      }
    }
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      return d.toISOString();
    }
  } catch (e) {
    // Fallback if parsing fails
  }
  return null;
};

export default function EKORunbookPage() {
  const [tasks, setTasks] = useState<EKOTask[]>([]);
  const [loading, setLoading] = useState(true);
  
  const [activeTab, setActiveTab] = useState<string>('ALL');

  const [newTitle, setNewTitle] = useState('');
  const [newAssigned, setNewAssigned] = useState('');
  const [newTeam, setNewTeam] = useState('');
  const [newScheduled, setNewScheduled] = useState('');
  const [newNotes, setNewNotes] = useState('');
  
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editTitle, setEditTitle] = useState('');
  const [editAssigned, setEditAssigned] = useState('');
  const [editTeam, setEditTeam] = useState('');
  const [editScheduled, setEditScheduled] = useState('');
  const [editNotes, setEditNotes] = useState('');

  const [uploadStatus, setUploadStatus] = useState<string | null>(null);

  const toDatetimeLocal = (isoStr: string | null) => {
    if (!isoStr) return '';
    const d = new Date(isoStr);
    d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
    return d.toISOString().slice(0, 16);
  };

  const fetchTasks = async () => {
    const { data, error } = await supabase
      .from('eko_tasks')
      .select('*');

    if (!error && data) {
      setTasks(data);
    }
    setLoading(false);
  };

  useEffect(() => {
    fetchTasks();

    const channel = supabase
      .channel('eko_tasks_realtime')
      .on(
        'postgres_changes',
        { event: '*', schema: 'public', table: 'eko_tasks' },
        () => fetchTasks()
      )
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
    };
  }, []);

  const handleClearAllTasks = async () => {
    if (confirm('Are you sure you want to delete ALL tasks in the tracker? This action cannot be undone.')) {
      setLoading(true);
      const { error } = await supabase.from('eko_tasks').delete().neq('id', '00000000-0000-0000-0000-000000000000');
      
      if (error) {
        setUploadStatus(`Error clearing tasks: ${error.message}`);
      } else {
        setTasks([]);
        setActiveTab('ALL');
        setUploadStatus('All tasks cleared successfully.');
      }
      setLoading(false);
    }
  };

  const toggleTask = async (id: string, currentStatus: boolean) => {
    setTasks((prev) =>
      prev.map((t) => (t.id === id ? { ...t, is_completed: !currentStatus } : t))
    );

    await supabase
      .from('eko_tasks')
      .update({ is_completed: !currentStatus })
      .eq('id', id);
  };

  const handleAddTask = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTitle) return; // Only title is explicitly required now

    const { error } = await supabase.from('eko_tasks').insert([
      {
        title: newTitle,
        subtask: 'Miscellaneous',
        assigned_to: newAssigned || 'Unassigned Person',
        team: newTeam || 'Unassigned Team',
        scheduled_at: newScheduled ? new Date(newScheduled).toISOString() : null, // Safely allow null
        notes: newNotes,
        is_completed: false,
      },
    ]);

    if (!error) {
      setNewTitle('');
      setNewAssigned('');
      setNewTeam('');
      setNewScheduled('');
      setNewNotes('');
    }
  };

  const startEditing = (task: EKOTask) => {
    setEditingId(task.id);
    setEditTitle(task.title);
    setEditAssigned(task.assigned_to);
    setEditTeam(task.team);
    setEditScheduled(task.scheduled_at ? toDatetimeLocal(task.scheduled_at) : '');
    setEditNotes(task.notes || '');
  };

  const saveAssignment = async (id: string) => {
    const updatedScheduledAt = editScheduled ? new Date(editScheduled).toISOString() : null;

    setTasks((prev) =>
      prev.map((t) =>
        t.id === id
          ? {
              ...t,
              title: editTitle,
              assigned_to: editAssigned,
              team: editTeam,
              scheduled_at: updatedScheduledAt,
              notes: editNotes,
            }
          : t
      )
    );

    await supabase
      .from('eko_tasks')
      .update({
        title: editTitle,
        assigned_to: editAssigned || 'Unassigned Person',
        team: editTeam || 'Unassigned Team',
        scheduled_at: updatedScheduledAt,
        notes: editNotes,
      })
      .eq('id', id);

    setEditingId(null);
  };

  const handleExcelImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const shouldReplace = tasks.length > 0 ? confirm('Do you want to REPLACE existing tasks? Click OK to replace all current tasks, or Cancel to ADD/APPEND to the existing list.') : false;

    setUploadStatus('Reading Excel file...');
    const reader = new FileReader();

    reader.onload = async (evt) => {
      try {
        const buffer = evt.target?.result;
        if (!buffer) throw new Error('File buffer is empty.');

        const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) throw new Error('No sheets found in workbook.');

        const firstSheet = workbook.Sheets[firstSheetName];
        const matrix: any[][] = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        let currentCategory = 'Topics Brainstorming';
        let colMap = { title: 0, person: -1, team: -1, date: -1, status: -1, notes: -1 };
        const mappedRows: any[] = [];

        for (let r = 0; r < matrix.length; r++) {
          const row = matrix[r];
          if (!Array.isArray(row) || row.length === 0) continue;

          const cellStrings = row.map((c) => (c !== null && c !== undefined ? String(c).trim() : ''));
          const nonEmpCells = cellStrings.filter((s) => s !== '');

          if (nonEmpCells.length === 0) continue; // Skip blank rows

          const rowCombinedLower = safeLower(cellStrings.join(' '));

          // 1. Column Headers Detection
          if (
            rowCombinedLower.includes('activity') ||
            rowCombinedLower.includes('assigned') ||
            rowCombinedLower.includes('things to do') ||
            rowCombinedLower.includes('owner')
          ) {
            cellStrings.forEach((val, idx) => {
              const l = safeLower(val);
              if (l.includes('activity') || l.includes('task') || l.includes('title') || l.includes('things to do')) colMap.title = idx;
              else if (l.includes('person') || l.includes('assigned') || l.includes('owner')) colMap.person = idx;
              else if (l.includes('team') || l.includes('department') || l.includes('group')) colMap.team = idx;
              else if (l.includes('date') || l.includes('time') || l.includes('schedule')) colMap.date = idx;
              else if (l.includes('status') || l.includes('done') || l.includes('completed') || l.includes('checkbox') || l.includes('state')) colMap.status = idx;
              else if (l.includes('note') || l.includes('comment') || l.includes('detail')) colMap.notes = idx;
            });
            continue;
          }

          // 2. Section Category Header Check
          const matchedCategory = detectCategoryHeader(cellStrings);
          if (matchedCategory) {
            currentCategory = matchedCategory;
            continue; 
          }

          // 3. Process Regular Task Row
          const getCell = (colIdx: number): string => (colIdx >= 0 && colIdx < cellStrings.length ? cellStrings[colIdx] : '');

          const rawTitle = getCell(colMap.title) || cellStrings[0] || '';
          if (!rawTitle) continue;

          const personVal = getCell(colMap.person);
          const teamVal = getCell(colMap.team);
          
          // Pull raw Date object from row if parsed by XLSX
          const dateVal = colMap.date >= 0 && colMap.date < row.length ? row[colMap.date] : '';
          
          const statusVal = getCell(colMap.status);
          const notesVal = getCell(colMap.notes);

          const scheduled_at = parseExcelDate(dateVal); // Extracts safely or returns null
          const is_completed = ['done', 'complete', 'completed', 'true', 'yes', 'x', '✓', '1'].includes(safeLower(statusVal));

          mappedRows.push({
            title: rawTitle,
            subtask: currentCategory,
            assigned_to: personVal || 'Unassigned Person',
            team: teamVal || 'Unassigned Team',
            scheduled_at,
            notes: notesVal,
            is_completed,
          });
        }

        if (mappedRows.length > 0) {
          if (shouldReplace) {
            await supabase.from('eko_tasks').delete().neq('id', '00000000-0000-0000-0000-000000000000');
          }

          const { error } = await supabase.from('eko_tasks').insert(mappedRows);
          if (error) {
            setUploadStatus(`Error importing: ${error.message}`);
          } else {
            setUploadStatus(`Successfully imported ${mappedRows.length} activities grouped under categories!`);
            fetchTasks();
          }
        } else {
          setUploadStatus('No valid tasks found in the sheet.');
        }
      } catch (err: any) {
        console.error('Excel Import Exception:', err);
        setUploadStatus(`Failed to parse file: ${err?.message || String(err)}`);
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const groupByCategory = (taskList: EKOTask[]) => {
    const groups: { [key: string]: EKOTask[] } = {};
    taskList.forEach((task) => {
      const cat = task.subtask || 'Miscellaneous';
      if (!groups[cat]) groups[cat] = [];
      groups[cat].push(task);
    });
    return groups;
  };

  const uniquePersons = Array.from(
    new Set(tasks.map((t) => t.assigned_to).filter((p) => p && p.trim() !== ''))
  ).sort();

  const filteredTasks = activeTab === 'ALL' 
    ? tasks 
    : tasks.filter((t) => safeLower(t.assigned_to) === safeLower(activeTab));

  const getDateKey = (dateStr: string | null) => {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  };

  // Safe Date Sorting (Places tasks with NO date at the bottom)
  const uncompletedTasks = filteredTasks
    .filter((t) => !t.is_completed)
    .sort((a, b) => {
      const timeA = a.scheduled_at ? new Date(a.scheduled_at).getTime() : Infinity;
      const timeB = b.scheduled_at ? new Date(b.scheduled_at).getTime() : Infinity;
      return timeA - timeB;
    });

  const uniqueUpcomingDates = Array.from(
    new Set(uncompletedTasks.map((t) => getDateKey(t.scheduled_at)).filter(Boolean)) // Filters out blank dates safely
  ).slice(0, 2);

  const upcomingTasksByDateGroup = uniqueUpcomingDates.map((dateKey) => ({
    dateKey,
    items: uncompletedTasks.filter((t) => getDateKey(t.scheduled_at) === dateKey),
  }));

  const isTaskInTopUpcomingDates = (scheduledAt: string | null) => {
    if (!scheduledAt) return false;
    return uniqueUpcomingDates.includes(getDateKey(scheduledAt));
  };

  const activeTasks = filteredTasks.filter((t) => !t.is_completed);
  const doneTasks = filteredTasks.filter((t) => t.is_completed);

  const activeGrouped = groupByCategory(activeTasks);
  const doneGrouped = groupByCategory(doneTasks);

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 p-4 sm:p-8">
      <div className="max-w-6xl mx-auto space-y-6">
        
        {/* Header */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between border-b border-slate-800 pb-6 gap-4">
          <div>
            <h1 className="text-3xl font-extrabold text-white tracking-tight">EKO Event Live Tracker</h1>
            <p className="text-slate-400 text-sm mt-1">Real-time team runbook & activity manager</p>
          </div>

          <div className="flex items-center gap-3">
            {tasks.length > 0 && (
              <button
                onClick={handleClearAllTasks}
                className="bg-red-950/60 hover:bg-red-900 text-red-300 font-medium px-3.5 py-2.5 rounded-lg border border-red-800 flex items-center gap-1.5 text-sm transition-colors"
                title="Wipe current list"
              >
                <Trash2 size={16} />
                <span>Clear All</span>
              </button>
            )}

            <label className="cursor-pointer bg-slate-800 hover:bg-slate-700 text-emerald-400 font-medium px-4 py-2.5 rounded-lg border border-slate-700 flex items-center justify-center gap-2 transition-colors">
              <Upload size={18} />
              <span>Upload Excel Sheet</span>
              <input type="file" accept=".xlsx, .xls, .csv" onChange={handleExcelImport} className="hidden" />
            </label>
          </div>
        </div>

        {uploadStatus && (
          <div className="p-3 rounded-lg bg-slate-800 text-slate-300 border border-slate-700 text-sm flex items-center justify-between">
            <div className="flex items-center gap-2">
              <AlertCircle size={16} className="text-emerald-400" />
              <span>{uploadStatus}</span>
            </div>
            <button onClick={() => setUploadStatus(null)} className="text-slate-400 hover:text-white">
              <X size={14} />
            </button>
          </div>
        )}

        {/* Dynamic Person Tabs Navigation */}
        <div className="flex items-center gap-2 overflow-x-auto pb-2 border-b border-slate-800 scrollbar-thin">
          <button
            onClick={() => setActiveTab('ALL')}
            className={`flex items-center gap-2 px-4 py-2 rounded-xl font-medium text-sm whitespace-nowrap transition-all ${
              activeTab === 'ALL'
                ? 'bg-emerald-600 text-white shadow-lg shadow-emerald-950/50'
                : 'bg-slate-800/80 text-slate-400 hover:text-slate-200 hover:bg-slate-800'
            }`}
          >
            <LayoutList size={16} />
            <span>Main Runbook (All)</span>
            <span className={`ml-1 text-xs px-2 py-0.5 rounded-full ${activeTab === 'ALL' ? 'bg-emerald-800 text-emerald-100' : 'bg-slate-900 text-slate-400'}`}>
              {tasks.length}
            </span>
          </button>

          {uniquePersons.map((person) => {
            const personTaskCount = tasks.filter((t) => safeLower(t.assigned_to) === safeLower(person)).length;
            const isSelected = safeLower(activeTab) === safeLower(person);

            return (
              <button
                key={person}
                onClick={() => setActiveTab(person)}
                className={`flex items-center gap-2 px-4 py-2 rounded-xl font-medium text-sm whitespace-nowrap transition-all ${
                  isSelected
                    ? 'bg-sky-600 text-white shadow-lg shadow-sky-950/50'
                    : 'bg-slate-800/80 text-slate-400 hover:text-slate-200 hover:bg-slate-800'
                }`}
              >
                <User size={15} />
                <span>{person}</span>
                <span className={`ml-1 text-xs px-2 py-0.5 rounded-full ${isSelected ? 'bg-sky-800 text-sky-100' : 'bg-slate-900 text-slate-400'}`}>
                  {personTaskCount}
                </span>
              </button>
            );
          })}
        </div>

        {/* UPCOMING TOP 2 DATES WIDGET */}
        <div className="bg-gradient-to-r from-orange-950/50 via-slate-800 to-slate-800 border-2 border-orange-500/60 rounded-2xl p-5 shadow-2xl shadow-orange-950/40">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-2">
              <Flame size={22} className="text-orange-400 animate-bounce" />
              <h2 className="text-lg font-bold text-orange-300">
                {activeTab === 'ALL' ? 'Upcoming Activities (Latest 2 Event Dates)' : `Upcoming Dates for ${activeTab}`}
              </h2>
            </div>
          </div>

          {upcomingTasksByDateGroup.length === 0 ? (
            <p className="text-slate-400 text-sm italic">No upcoming pending activities found for this view!</p>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
              {upcomingTasksByDateGroup.map((group, idx) => (
                <div key={group.dateKey} className="bg-slate-900/90 border border-orange-500/40 rounded-xl p-4 shadow-md space-y-3">
                  <div className="flex items-center justify-between border-b border-orange-500/20 pb-2">
                    <span className="flex items-center gap-1.5 text-xs font-extrabold uppercase tracking-wider text-orange-400 bg-orange-500/20 px-2.5 py-1 rounded-md border border-orange-500/30">
                      <Calendar size={13} /> Date #{idx + 1}: {group.dateKey}
                    </span>
                    <span className="text-xs font-semibold text-orange-300 bg-slate-800 px-2 py-0.5 rounded border border-slate-700">
                      {group.items.length} {group.items.length === 1 ? 'Task' : 'Tasks'}
                    </span>
                  </div>

                  <div className="space-y-2.5 max-h-60 overflow-y-auto pr-1">
                    {group.items.map((task) => (
                      <div key={task.id} className="bg-slate-950/80 border border-slate-800 p-2.5 rounded-lg space-y-1">
                        <div className="flex items-start justify-between gap-2">
                          <h4 className="font-semibold text-orange-100 text-xs">{task.title}</h4>
                          {task.scheduled_at && (
                            <span className="text-[10px] text-orange-400 font-mono shrink-0 bg-orange-950/50 px-1.5 py-0.5 rounded border border-orange-900/50">
                              {new Date(task.scheduled_at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                            </span>
                          )}
                        </div>

                        {task.subtask && (
                          <div className="flex items-center gap-1 text-[11px] text-purple-300 font-medium bg-purple-950/40 px-2 py-0.5 rounded w-fit border border-purple-800/40">
                            <Tag size={10} />
                            <span>{task.subtask}</span>
                          </div>
                        )}

                        <div className="flex items-center gap-2 text-[10px] text-slate-400 pt-0.5">
                          <span className="text-emerald-400 font-medium"><Users size={10} className="inline mr-0.5" />{task.team}</span>
                          <span>•</span>
                          <span className="text-sky-400 font-medium"><User size={10} className="inline mr-0.5" />{task.assigned_to}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Add Activity Form */}
        <form onSubmit={handleAddTask} className="bg-slate-800 border border-slate-700 p-4 rounded-xl grid grid-cols-1 md:grid-cols-3 gap-3 shadow-lg">
          <input
            type="text"
            placeholder="Task Description / Topic Activity..."
            value={newTitle}
            onChange={(e) => setNewTitle(e.target.value)}
            className="bg-slate-900 border border-slate-700 rounded-lg p-2.5 text-sm md:col-span-2 text-white focus:outline-none focus:ring-2 focus:ring-emerald-500"
            required
          />
          {/* Due date is optional now (no required flag) */}
          <input
            type="datetime-local"
            value={newScheduled}
            onChange={(e) => setNewScheduled(e.target.value)}
            className="bg-slate-900 border border-slate-700 rounded-lg p-2.5 text-sm text-white focus:outline-none focus:ring-2 focus:ring-emerald-500"
          />
          <input
            type="text"
            placeholder="Team (e.g. AV / Tech)"
            value={newTeam}
            onChange={(e) => setNewTeam(e.target.value)}
            className="bg-slate-900 border border-slate-700 rounded-lg p-2.5 text-sm text-white focus:outline-none focus:ring-2 focus:ring-emerald-500"
          />
          <input
            type="text"
            placeholder="Person (e.g. Alex)"
            value={newAssigned}
            onChange={(e) => setNewAssigned(e.target.value)}
            className="bg-slate-900 border border-slate-700 rounded-lg p-2.5 text-sm text-white focus:outline-none focus:ring-2 focus:ring-emerald-500"
          />
          <input
            type="text"
            placeholder="Notes / Additional details..."
            value={newNotes}
            onChange={(e) => setNewNotes(e.target.value)}
            className="bg-slate-900 border border-slate-700 rounded-lg p-2.5 text-sm text-white focus:outline-none focus:ring-2 focus:ring-emerald-500"
          />
          <button
            type="submit"
            className="md:col-span-3 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold py-2.5 rounded-lg flex items-center justify-center gap-2 transition-colors"
          >
            <Plus size={18} /> Add Activity
          </button>
        </form>

        {/* TWO-COLUMN BOARD VIEW split into visual Category Sections */}
        {loading ? (
          <p className="text-slate-500 text-center py-10">Loading runbook activities...</p>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            
            {/* COLUMN 1: PENDING TASKS GROUPED BY CATEGORY */}
            <div className="space-y-6">
              <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                <h3 className="font-bold text-slate-200 flex items-center gap-2 text-base">
                  <Clock size={18} className="text-emerald-400" />
                  To Do / Pending ({activeTab === 'ALL' ? 'All' : activeTab})
                </h3>
                <span className="bg-slate-800 text-slate-300 text-xs px-2.5 py-1 rounded-full font-semibold">
                  {activeTasks.length}
                </span>
              </div>

              {Object.keys(activeGrouped).length === 0 ? (
                <div className="text-center py-8 border border-dashed border-slate-800 rounded-xl text-slate-500 text-sm">
                  No pending activities for this view!
                </div>
              ) : (
                Object.entries(activeGrouped).map(([categoryName, groupTasks]) => (
                  <div key={categoryName} className="space-y-3 bg-slate-950/40 p-3.5 rounded-2xl border border-slate-800">
                    <div className="flex items-center justify-between bg-purple-950/60 border border-purple-800/50 p-2.5 rounded-xl">
                      <div className="flex items-center gap-2">
                        <Folder size={16} className="text-purple-400" />
                        <h4 className="font-bold text-xs uppercase tracking-wider text-purple-200">{categoryName}</h4>
                      </div>
                      <span className="bg-purple-900/80 text-purple-300 text-[11px] px-2 py-0.5 rounded-full font-semibold">
                        {groupTasks.length}
                      </span>
                    </div>

                    <div className="space-y-3">
                      {groupTasks.map((task) => {
                        const isUpcomingDateTask = isTaskInTopUpcomingDates(task.scheduled_at);

                        return (
                          <div
                            key={task.id}
                            className={`p-4 rounded-xl border transition-all flex items-start justify-between gap-3 ${
                              isUpcomingDateTask
                                ? 'bg-gradient-to-r from-orange-950/40 via-slate-800 to-slate-800 border-orange-500/80 shadow-lg shadow-orange-950/30 ring-1 ring-orange-500/50'
                                : 'bg-slate-800/80 border-slate-700 hover:border-slate-600'
                            }`}
                          >
                            <div className="flex items-start gap-3 w-full">
                              <button
                                onClick={() => toggleTask(task.id, task.is_completed)}
                                className="mt-1 text-slate-400 hover:text-emerald-400 transition-colors shrink-0"
                              >
                                <Circle size={22} />
                              </button>

                              <div className="space-y-2 w-full">
                                <div className="flex items-center justify-between gap-2">
                                  <h5 className={`font-semibold text-sm ${isUpcomingDateTask ? 'text-orange-100' : 'text-slate-100'}`}>
                                    {task.title}
                                  </h5>

                                  {isUpcomingDateTask && (
                                    <span className="text-[10px] font-bold text-orange-300 bg-orange-500/20 px-2 py-0.5 rounded border border-orange-500/40 flex items-center gap-1 shrink-0">
                                      <Flame size={10} /> Upcoming Date
                                    </span>
                                  )}
                                </div>

                                {editingId === task.id ? (
                                  <div className="bg-slate-900 border border-slate-700 p-3 rounded-lg space-y-2 text-xs">
                                    <div>
                                      <label className="text-slate-400 block mb-1">Task Title</label>
                                      <input
                                        type="text"
                                        value={editTitle}
                                        onChange={(e) => setEditTitle(e.target.value)}
                                        className="w-full bg-slate-800 border border-slate-600 p-1.5 rounded text-white focus:outline-none"
                                      />
                                    </div>

                                    <div className="grid grid-cols-2 gap-2">
                                      <div>
                                        <label className="text-slate-400 block mb-1">Team</label>
                                        <input
                                          type="text"
                                          value={editTeam}
                                          onChange={(e) => setEditTeam(e.target.value)}
                                          className="w-full bg-slate-800 border border-slate-600 p-1.5 rounded text-white focus:outline-none"
                                        />
                                      </div>
                                      <div>
                                        <label className="text-slate-400 block mb-1">Person</label>
                                        <input
                                          type="text"
                                          value={editAssigned}
                                          onChange={(e) => setEditAssigned(e.target.value)}
                                          className="w-full bg-slate-800 border border-slate-600 p-1.5 rounded text-white focus:outline-none"
                                        />
                                      </div>
                                    </div>

                                    <div>
                                      <label className="text-slate-400 block mb-1">Due Date & Time</label>
                                      <input
                                        type="datetime-local"
                                        value={editScheduled}
                                        onChange={(e) => setEditScheduled(e.target.value)}
                                        className="w-full bg-slate-800 border border-slate-600 p-1.5 rounded text-white focus:outline-none"
                                      />
                                    </div>

                                    <div>
                                      <label className="text-slate-400 block mb-1">Notes</label>
                                      <textarea
                                        value={editNotes}
                                        onChange={(e) => setEditNotes(e.target.value)}
                                        rows={2}
                                        className="w-full bg-slate-800 border border-slate-600 p-1.5 rounded text-white focus:outline-none"
                                      />
                                    </div>

                                    <div className="flex items-center gap-2 pt-1">
                                      <button
                                        onClick={() => saveAssignment(task.id)}
                                        className="bg-emerald-600 text-white px-3 py-1 rounded hover:bg-emerald-500 flex items-center gap-1"
                                      >
                                        <Save size={12} /> Save
                                      </button>
                                      <button
                                        onClick={() => setEditingId(null)}
                                        className="bg-slate-700 text-slate-300 px-3 py-1 rounded hover:bg-slate-600 flex items-center gap-1"
                                      >
                                        <X size={12} /> Cancel
                                      </button>
                                    </div>
                                  </div>
                                ) : (
                                  <>
                                    {task.notes && (
                                      <p className="text-xs text-slate-300 bg-slate-900/80 border border-slate-800 p-2 rounded-lg flex items-start gap-1.5">
                                        <FileText size={13} className="text-amber-400 shrink-0 mt-0.5" />
                                        <span>{task.notes}</span>
                                      </p>
                                    )}

                                    <div className="flex flex-wrap items-center gap-2 text-xs text-slate-400 pt-1">
                                      {/* Only render Date chip if date exists */}
                                      {task.scheduled_at && (
                                        <span className={`flex items-center gap-1 px-2 py-0.5 rounded border ${isUpcomingDateTask ? 'bg-orange-950/80 border-orange-500/50 text-orange-300 font-semibold' : 'bg-slate-900 border-slate-800'}`}>
                                          <Clock size={12} />
                                          {new Date(task.scheduled_at).toLocaleString([], {
                                            dateStyle: 'short',
                                            timeStyle: 'short',
                                          })}
                                        </span>
                                      )}

                                      <span className="flex items-center gap-1 bg-slate-900 px-2 py-0.5 rounded border border-slate-800 text-emerald-300">
                                        <Users size={12} /> {task.team}
                                      </span>

                                      <span className="flex items-center gap-1 bg-slate-900 px-2 py-0.5 rounded border border-slate-800 text-sky-300">
                                        <User size={12} /> {task.assigned_to}
                                      </span>

                                      <button
                                        onClick={() => startEditing(task)}
                                        className="text-slate-500 hover:text-slate-300 p-0.5 flex items-center gap-1 text-[11px] bg-slate-900 px-2 py-0.5 rounded border border-slate-800"
                                      >
                                        <Edit3 size={11} /> Edit
                                      </button>
                                    </div>
                                  </>
                                )}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ))
              )}
            </div>

            {/* COLUMN 2: COMPLETED TASKS GROUPED BY CATEGORY */}
            <div className="space-y-6">
              <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                <h3 className="font-bold text-slate-200 flex items-center gap-2 text-base">
                  <CheckCheck size={18} className="text-emerald-400" />
                  Completed Column ({activeTab === 'ALL' ? 'All' : activeTab})
                </h3>
                <span className="bg-emerald-950 text-emerald-400 border border-emerald-800 text-xs px-2.5 py-1 rounded-full font-semibold">
                  {doneTasks.length}
                </span>
              </div>

              {Object.keys(doneGrouped).length === 0 ? (
                <div className="text-center py-8 border border-dashed border-slate-800 rounded-xl text-slate-500 text-sm">
                  Completed tasks will move here automatically!
                </div>
              ) : (
                Object.entries(doneGrouped).map(([categoryName, groupTasks]) => (
                  <div key={categoryName} className="space-y-3 bg-slate-950/40 p-3.5 rounded-2xl border border-slate-800 opacity-80">
                    <div className="flex items-center justify-between bg-emerald-950/40 border border-emerald-800/40 p-2.5 rounded-xl">
                      <div className="flex items-center gap-2">
                        <Folder size={16} className="text-emerald-400" />
                        <h4 className="font-bold text-xs uppercase tracking-wider text-emerald-200">{categoryName}</h4>
                      </div>
                      <span className="bg-emerald-900/80 text-emerald-300 text-[11px] px-2 py-0.5 rounded-full font-semibold">
                        {groupTasks.length}
                      </span>
                    </div>

                    <div className="space-y-2">
                      {groupTasks.map((task) => (
                        <div
                          key={task.id}
                          className="p-3.5 rounded-xl border bg-slate-950/60 border-slate-800/80 flex items-start justify-between gap-3"
                        >
                          <div className="flex items-start gap-3 w-full">
                            <button
                              onClick={() => toggleTask(task.id, task.is_completed)}
                              className="mt-1 text-emerald-400 hover:text-slate-400 transition-colors shrink-0"
                            >
                              <CheckCircle2 size={20} />
                            </button>

                            <div className="space-y-1 w-full">
                              <h5 className="font-semibold text-slate-400 line-through text-xs">{task.title}</h5>

                              {task.notes && (
                                <p className="text-xs text-slate-500 italic bg-slate-900/50 p-1.5 rounded">
                                  "{task.notes}"
                                </p>
                              )}

                              <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500 pt-1">
                                {/* Only render Date chip if date exists */}
                                {task.scheduled_at && (
                                  <span className="flex items-center gap-1 bg-slate-900 px-2 py-0.5 rounded border border-slate-800">
                                    <Clock size={11} />
                                    {new Date(task.scheduled_at).toLocaleString([], {
                                      dateStyle: 'short',
                                      timeStyle: 'short',
                                    })}
                                  </span>
                                )}

                                <span className="flex items-center gap-1 bg-slate-900 px-2 py-0.5 rounded border border-slate-800 text-slate-400">
                                  <Users size={11} /> {task.team}
                                </span>

                                <span className="flex items-center gap-1 bg-slate-900 px-2 py-0.5 rounded border border-slate-800 text-slate-400">
                                  <User size={11} /> {task.assigned_to}
                                </span>
                              </div>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                ))
              )}
            </div>

          </div>
        )}

      </div>
    </div>
  );
}
