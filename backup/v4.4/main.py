import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, time, json, queue
from pynput import mouse, keyboard

root = tk.Tk()
root.title('Desktop AI Agent v4.4 — MCP Edition')
root.geometry('620x720')
root.configure(bg='#eef2f7')
root.attributes('-topmost', False)
controller = mouse.Controller()
commands = queue.Queue()
lock = threading.Lock()
stop = threading.Event()
pause = threading.Event()
events = []
recording = False
started = 0
worker = None
loops = 0
status = tk.StringVar(value='Ready')
counts = tk.StringVar(value='0 events     |     0 loops')
speed = tk.DoubleVar(value=1.0)
repeat = tk.BooleanVar(value=True)

def capture(action, *args):
    with lock:
        if recording:
            events.append([time.monotonic()-started, action, *args])

def click(x,y,button,pressed):
    capture('click',x,y,button.name,pressed)

listener = mouse.Listener(on_move=lambda x,y: capture('move',x,y), on_click=click,
                          on_scroll=lambda x,y,dx,dy: capture('scroll',x,y,dx,dy))
listener.start()

def busy():
    return worker is not None and worker.is_alive()

def record():
    global recording, started, loops
    if busy() or recording: return
    root.iconify()
    status.set('Starting recording...')
    def begin():
        global recording, started, loops
        if status.get() != 'Starting recording...': return
        with lock:
            events.clear(); started=time.monotonic(); recording=True; loops=0
        status.set('Recording')
    root.after(500,begin)

def finish():
    global recording
    with lock: recording=False
    if not busy():
        status.set('Recording stopped'); root.deiconify()

def wait_delay(delay):
    while delay > 0 or pause.is_set():
        if stop.is_set(): return False
        if pause.is_set():
            stop.wait(.03); continue
        step=min(.02,delay)
        before=time.monotonic()
        if stop.wait(step): return False
        if not pause.is_set(): delay-=time.monotonic()-before
    return not stop.is_set()

def playback(data, rate, looping):
    global loops
    held=set()
    try:
        while not stop.is_set():
            previous=0
            for event in data:
                t,action,*a=event
                if not wait_delay(max(0,t-previous)/rate): return
                previous=t
                controller.position=tuple(a[:2])
                if action=='click':
                    button=getattr(mouse.Button,a[2])
                    if a[3]: controller.press(button); held.add(button)
                    else: controller.release(button); held.discard(button)
                elif action=='scroll': controller.scroll(a[2],a[3])
            for button in held: controller.release(button)
            held.clear()
            loops+=1
            if not looping or not wait_delay(.2): break
    except Exception as exc:
        commands.put(('error',str(exc)))
    finally:
        for button in held: controller.release(button)
        commands.put(('done',None))

def play():
    global worker, loops
    if recording or busy() or status.get()=='Starting recording...': return
    with lock: data=[e[:] for e in events]
    if not data: return
    stop.clear(); pause.clear(); loops=0
    status.set('Running'); root.iconify()
    worker=threading.Thread(target=playback,args=(data,speed.get(),repeat.get()),daemon=True)
    worker.start()

def toggle():
    if not busy() or stop.is_set(): return
    if pause.is_set(): pause.clear(); status.set('Running'); ai_status.set('Requesting next action...'); root.iconify()
    else: pause.set(); status.set('Paused'); ai_status.set('Paused')

def halt():
    stop.set(); pause.clear()
    if busy(): status.set('Stopping...'); ai_status.set('Stopping...')

def clear():
    global loops
    if recording or busy(): return
    with lock: events.clear()
    loops=0; status.set('Ready')

def save():
    if recording or busy(): return
    path=filedialog.asksaveasfilename(defaultextension='.json',filetypes=[('JSON','*.json')])
    if path:
        try:
            with open(path,'w') as f: json.dump(events,f)
        except Exception as exc: messagebox.showerror('Save error',str(exc))

def load():
    if recording or busy(): return
    path=filedialog.askopenfilename(filetypes=[('JSON','*.json')])
    if not path: return
    try:
        with open(path) as f: data=json.load(f)
        if not isinstance(data,list): raise ValueError('Invalid recording')
        previous=0
        for e in data:
            if not isinstance(e,list) or len(e)<4 or not isinstance(e[0],(int,float)) or not previous<=e[0]<float('inf'): raise ValueError('Invalid event')
            previous=e[0]
            if e[1] not in ('move','click','scroll') or len(e)!={'move':4,'click':6,'scroll':6}[e[1]]: raise ValueError('Invalid action')
            if not all(type(v) is int for v in e[2:4]): raise ValueError('Invalid coordinates')
            if e[1]=='click' and (e[4] not in ('left','right','middle','x1','x2') or type(e[5]) is not bool): raise ValueError('Invalid click')
            if e[1]=='scroll' and not all(type(v) is int for v in e[4:]): raise ValueError('Invalid scroll')
        with lock: events[:]=data
        status.set('Recording loaded')
    except Exception as exc: messagebox.showerror('Load error',str(exc))

closing=False
def close():
    global closing, recording
    closing=True; recording=False; halt(); listener.stop(); keys.stop()
    if not busy(): root.destroy()

def poll():
    while True:
        try: action,arg=commands.get_nowait()
        except queue.Empty: break
        if action=='done':
            status.set('Ready')
            if not closing: root.deiconify()
        elif action=='error':
            if not closing: messagebox.showerror('Playback error',arg)
        elif not closing: action()
    if closing and not busy(): root.destroy(); return
    counts.set(f'{len(events):,} events     |     {loops:,} loops')
    root.after(50,poll)

# AI agent: one screenshot and one proposed action per step.
import base64, io, ctypes, sys
from openai import OpenAI, RateLimitError
from pathlib import Path
import os
import task_memory
import win32crypt
import win32cred
import pyautogui as pg
pg.FAILSAFE = True

base = tk.StringVar(value='https://ai-api.orangevpn.com.ph/v1')
model = tk.StringVar(value='anthropic/claude-haiku-5.5')
api_key = tk.StringVar()
ai_status = tk.StringVar(value='Ready')
limit = tk.IntVar(value=20)
approve = tk.BooleanVar(value=True)

# Store non-secret preferences in AppData and the key using Windows DPAPI encryption.
settings_path = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'DesktopAI' / 'api_settings.json'
credential_target = 'DesktopAI.v4.APIKey'
key_path = settings_path.with_name('api_key.dpapi')
remember_key = tk.BooleanVar(value=False)
save_status = tk.StringVar(value='Settings not saved yet.')
connection_status = tk.StringVar(value='Not tested')
connection_detail = tk.StringVar(value='')

def load_api_settings():
    try:
        if not settings_path.exists(): return
        data = json.loads(settings_path.read_text(encoding='utf-8'))
        base.set(data.get('base_url', base.get()))
        model.set(data.get('model', model.get()))
        remember_key.set(bool(data.get('remember_key', False)))
        if remember_key.get():
            if key_path.exists():
                _, raw = win32crypt.CryptUnprotectData(key_path.read_bytes(), None, None, None, 1)
                api_key.set(raw.decode('utf-8'))
            else:
                cred = win32cred.CredRead(credential_target, win32cred.CRED_TYPE_GENERIC, 0)
                blob = cred['CredentialBlob']
                api_key.set(blob.decode('utf-16-le') if isinstance(blob, bytes) else blob)
        save_status.set('Saved settings loaded.')
    except Exception:
        save_status.set('Could not load all saved settings. Enter your key again.')

def save_api_settings():
    stage = 'validating settings'
    try:
        url, key, name = settings()
        settings_path.parent.mkdir(parents=True, exist_ok=True)
        stage = 'encrypting API key'
        if remember_key.get():
            encrypted = win32crypt.CryptProtectData(key.encode('utf-8'), 'DesktopAI API key', None, None, None, 1)
            stage = 'writing encrypted key'
            key_temp = key_path.with_suffix('.tmp')
            key_temp.write_bytes(encrypted)
            key_temp.replace(key_path)
        else:
            stage = 'removing saved key'
            key_path.unlink(missing_ok=True)
            try: win32cred.CredDelete(credential_target, win32cred.CRED_TYPE_GENERIC, 0)
            except Exception as exc:
                if getattr(exc, 'winerror', None) != 1168: raise
        stage = 'writing preferences'
        settings_path.parent.mkdir(parents=True, exist_ok=True)
        temp = settings_path.with_suffix('.tmp')
        temp.write_text(json.dumps({'base_url': url, 'model': name,
            'remember_key': bool(remember_key.get())}, indent=2), encoding='utf-8')
        temp.replace(settings_path)
        save_status.set('Settings saved successfully.')
    except Exception as exc:
        # Do not display exception arguments that could contain the API key.
        code = getattr(exc, 'winerror', None)
        if code is None and isinstance(exc, OSError): code = exc.errno
        detail = type(exc).__name__ + (f' (code {code})' if code is not None else '')
        save_status.set('Save failed while ' + stage + ': ' + detail)

def show_connection(state, detail, color):
    def update():
        connection_status.set(state)
        connection_detail.set(detail)
        connection_label.configure(foreground=color)
    commands.put((update, None))

def log_line(text):
    def update():
        if text.startswith('Step '): ai_status.set('Analyzing screen...')
        elif text.startswith(('API cooldown:', 'Rate limit reached', 'API timeout or connection failure')): ai_status.set('Waiting for API cooldown...')
        elif text.startswith('Waiting for AI response'): ai_status.set('Waiting for AI response...')
        elif text.startswith('Continuing:'): ai_status.set('Requesting next action...')
        elif text.startswith('AI response'): ai_status.set('Checking AI response...')
        elif text.split(' ',1)[0] in ('click','move','type_text','hotkey','press','scroll','list_windows','focus_window','excel_write','screen_size','draw_path'): ai_status.set('Executing / awaiting approval: '+text.split(' ',1)[0])
        log.insert('end', text+'\n'); log.see('end')
    commands.put((update,None))

def settings():
    url=base.get().strip(); key=api_key.get().strip(); name=model.get().strip()
    if not url.startswith('https://') or not key or not name:
        raise ValueError('Enter an HTTPS base URL, API key and model ID.')
    return url,key,name

def confirm_action(a):
    result=queue.Queue()
    def dialog():
        if stop.is_set() or closing: result.put(False); return
        result.put(messagebox.askyesno('Approve AI action',json.dumps(a,ensure_ascii=False,indent=2),parent=root))
    commands.put((dialog,None))
    while not stop.is_set():
        try: return result.get(timeout=.1)
        except queue.Empty: pass
    return False

def ai_run(config,task,max_steps,approval):
    import asyncio
    from agent_engine import run_agent
    def confirm(a):
        foreground=ctypes.windll.user32.GetForegroundWindow()
        if not confirm_action(a): return False
        if ctypes.windll.user32.GetForegroundWindow()!=foreground:
            ctypes.windll.user32.SetForegroundWindow(foreground)
            if not wait_delay(.25): return False
        return ctypes.windll.user32.GetForegroundWindow()==foreground
    outcome="Stopped"
    try:
        if wait_delay(.8):
            outcome=asyncio.run(run_agent(config,task,max_steps,approval,log_line,stop.is_set,wait_delay,confirm)) or "Incomplete"
    except Exception as exc:
        import traceback
        outcome='Error — see action log'
        log_line('AI error:\n'+traceback.format_exc())
    finally:
        final_status='Stopped' if stop.is_set() else outcome
        try: task_memory.append('task outcome',final_status)
        except Exception as exc: log_line('Memory save failed: '+type(exc).__name__)
        commands.put((lambda: (ai_status.set(final_status),refresh_memory()),None))
        commands.put(('done',None))

def refresh_memory():
    try: memory_status.set(f'Memory: {len(task_memory.load())} saved entries')
    except Exception: memory_status.set('Memory unavailable — reset to repair')

def reset_tasks():
    if busy():
        messagebox.showinfo('Reset tasks','Stop the active task and wait until it finishes first.')
        return
    if not messagebox.askyesno('Reset all tasks','Clear saved task history, instructions and action log? API settings and recordings will remain.'): return
    try: task_memory.reset()
    except OSError as exc:
        messagebox.showerror('Reset failed',str(exc)); return
    log.delete('1.0','end'); task_box.delete('1.0','end')
    ai_status.set('Ready'); refresh_memory()

def continue_task():
    if busy(): return
    try:
        if not task_memory.load():
            messagebox.showinfo('Continue','No previous task saved.'); return
    except Exception:
        messagebox.showerror('Memory','Cannot read task history.'); return
    task_box.delete('1.0','end')
    task_box.insert('1.0','Continue the previous task from its last verified progress. Check the current desktop first.')
    start_ai()

def start_ai():
    global worker
    if busy() or recording or status.get()=='Starting recording...': return
    try:
        config=settings(); task=task_box.get('1.0','end').strip(); steps=int(limit.get())
        if not task or not 1<=steps<=100: raise ValueError('Enter a task and step limit (1–100).')
    except Exception as exc: messagebox.showerror('Settings',str(exc)); return
    if not messagebox.askyesno('Screen sharing','Screenshots of your primary display will be sent to your configured API provider. Close private windows first. Continue?'): return
    approval = bool(approve.get())
    if not approval and not messagebox.askyesno('Automatic mode', 'Desktop actions will execute without individual approval. Keep F10 ready to stop. Continue?'):
        return
    log_line('Manual approval: '+('ON' if approval else 'OFF'))
    stop.clear(); pause.clear(); status.set('AI running'); ai_status.set('Starting AI...'); root.iconify()
    worker=threading.Thread(target=ai_run,args=(config,task,steps,approval),daemon=True); worker.start()

def test_api():
    global worker
    if busy() or recording: return
    try: config=settings()
    except Exception as exc:
        show_connection('Invalid settings', str(exc), '#b91c1c'); return
    stop.clear(); pause.clear(); status.set('Testing API')
    test_button.configure(state='disabled')
    for entry in settings_entries: entry.configure(state='disabled')
    show_connection('Testing connection...', 'Checking text API and native tool calling.', '#a16207')
    def run():
        text_ok = False
        try:
            with OpenAI(base_url=config[0],api_key=config[1],timeout=25,max_retries=0) as client:
                response=client.chat.completions.create(model=config[2],messages=[{'role':'user','content':'Reply OK'}])
            text_ok = True
            log_line('Text API test: OK')
            import asyncio
            from agent_engine import test_tools
            passed = asyncio.run(test_tools(config))
            log_line('Native tool calling: '+('PASS' if passed else 'NOT CONFIRMED'))
            show_connection('Connected' if passed else 'Partially connected',
                'Text API: OK | Native tools: '+('PASS' if passed else 'NOT CONFIRMED')+' | Vision: not tested',
                '#15803d' if passed else '#a16207')
        except RateLimitError:
            show_connection('Rate limited (429)', ('Text API: OK. ' if text_ok else '')+'Wait about 60–65 seconds and retry.', '#b91c1c')
            log_line('API test: rate limited (429)')
        except Exception as exc:
            detail = ('Text API: OK. Tool test failed. ' if text_ok else '') + type(exc).__name__
            show_connection('Test failed', detail, '#b91c1c')
            log_line('API test failed: '+detail)
        finally:
            def restore():
                test_button.configure(state='normal')
                for entry in settings_entries: entry.configure(state='normal')
            commands.put((restore,None))
            commands.put(('done',None))
    worker=threading.Thread(target=run,daemon=True); worker.start()


style=ttk.Style(root)
style.theme_use('clam')
style.configure('TButton',padding=9,font=('Segoe UI',10))
style.configure('TFrame',background='#eef2f7')
style.configure('TLabel',background='#eef2f7',font=('Segoe UI',10))
style.configure('TCheckbutton',background='#eef2f7')
notebook=ttk.Notebook(root); notebook.pack(fill='both',expand=True,padx=10,pady=10)
frame=ttk.Frame(notebook,padding=20); notebook.add(frame,text='Recorder')
ttk.Label(frame,text='Auto Clicker Pro',font=('Segoe UI',22,'bold')).pack(anchor='w')
ttk.Label(frame,text='Mouse recorder & automation').pack(anchor='w')
ttk.Label(frame,textvariable=status,font=('Segoe UI',12,'bold')).pack(pady=(18,4))
ttk.Label(frame,textvariable=counts).pack(pady=(0,12))
ttk.Button(frame,text='Start Recording  [F6]',command=record).pack(fill='x')
buttons=ttk.Frame(frame); buttons.pack(fill='x',pady=5)
for i,(label,fn) in enumerate([('Stop Recording [F7]',finish),('Play Loop [F8]',play),('Pause / Resume [F9]',toggle),('Stop Playback [F10]',halt)]):
    ttk.Button(buttons,text=label,command=fn).grid(row=i//2,column=i%2,sticky='ew',padx=2,pady=2)
for i in range(2): buttons.columnconfigure(i,weight=1)
ttk.Label(frame,text='Playback speed (applies on next play)').pack(anchor='w',pady=(10,0))
rate_label=ttk.Label(frame,text='1.00x'); rate_label.pack(anchor='e')
ttk.Scale(frame,from_=.25,to=4,variable=speed,command=lambda v: rate_label.configure(text=f'{float(v):.2f}x')).pack(fill='x')
ttk.Checkbutton(frame,text='Repeat continuously',variable=repeat).pack(anchor='w',pady=8)
files=ttk.Frame(frame); files.pack(fill='x')
for label,fn in [('Save',save),('Load',load),('Clear',clear)]: ttk.Button(files,text=label,command=fn).pack(side='left',expand=True,fill='x',padx=2)
ttk.Label(frame,text='F10 stops playback • F12 exits').pack(pady=12)
agent=ttk.Frame(notebook,padding=16); notebook.add(agent,text='AI Agent')
ttk.Label(agent,text='AI Desktop Agent',font=('Segoe UI',20,'bold')).pack(anchor='w')
ttk.Label(agent,textvariable=ai_status,font=('Segoe UI',11,'bold'),foreground='#2563eb').pack(anchor='w')
activity=ttk.Progressbar(agent,mode='indeterminate')
activity.pack(fill='x',pady=(3,5))
def update_activity(*args):
    active=ai_status.get().startswith(('Starting','Analyzing','Waiting','Requesting','Checking','Executing'))
    if active: activity.start(15)
    else: activity.stop()
ai_status.trace_add('write',update_activity)
ttk.Label(agent,text='Task instructions').pack(anchor='w',pady=(10,0))
task_box=tk.Text(agent,height=4,wrap='word'); task_box.pack(fill='x')
ttk.Checkbutton(agent,text='Approve each action (recommended)',variable=approve).pack(anchor='w',pady=6)
line=ttk.Frame(agent); line.pack(fill='x')
ttk.Label(line,text='Maximum steps').pack(side='left'); ttk.Spinbox(line,from_=1,to=100,textvariable=limit,width=6).pack(side='left',padx=8)
line=ttk.Frame(agent); line.pack(fill='x',pady=8)
for label,fn in [('Start AI',start_ai),('Pause / Resume F9',toggle),('Stop F10',halt)]: ttk.Button(line,text=label,command=fn).pack(side='left',expand=True,fill='x',padx=2)
memory_status=tk.StringVar()
memline=ttk.Frame(agent); memline.pack(fill='x',pady=(0,6))
ttk.Label(memline,textvariable=memory_status).pack(side='left')
ttk.Button(memline,text='Reset All Tasks',command=reset_tasks).pack(side='right')
ttk.Button(memline,text='Continue Task',command=continue_task).pack(side='right',padx=4)
refresh_memory()
ttk.Label(agent,text='Action log').pack(anchor='w')
log=tk.Text(agent,height=12,wrap='word'); log.pack(fill='both',expand=True)
tools_tab=ttk.Frame(notebook,padding=16); notebook.add(tools_tab,text='MCP Tools')
ttk.Label(tools_tab,text='Local MCP server (stdio)',font=('Segoe UI',18,'bold')).pack(anchor='w')
ttk.Label(tools_tab,text='Starts automatically for each AI task. No network listener.',wraplength=530).pack(anchor='w',pady=10)
for name in ['screenshot','screen_size','click','move','type_text','hotkey','press','scroll','list_windows','focus_window','excel_write']:
    ttk.Label(tools_tab,text=name).pack(anchor='w')
settings_tab=ttk.Frame(notebook,padding=16); notebook.add(settings_tab,text='API Settings')
settings_entries=[]
for label,var,secret in [('Base URL',base,False),('API Key',api_key,True),('Model ID',model,False)]:
    ttk.Label(settings_tab,text=label).pack(anchor='w',pady=(12,3)); entry=ttk.Entry(settings_tab,textvariable=var,show='*' if secret else ''); entry.pack(fill='x'); settings_entries.append(entry)
ttk.Checkbutton(settings_tab,text='Remember API key securely on this PC',variable=remember_key).pack(anchor='w',pady=(14,4))
ttk.Button(settings_tab,text='Save Settings',command=save_api_settings).pack(fill='x',pady=4)
ttk.Label(settings_tab,textvariable=save_status,wraplength=530).pack(anchor='w',pady=4)
test_button=ttk.Button(settings_tab,text='Test API and native tool calling',command=test_api)
test_button.pack(fill='x',pady=12)
ttk.Separator(settings_tab).pack(fill='x',pady=8)
ttk.Label(settings_tab,text='Connection status',font=('Segoe UI',11,'bold')).pack(anchor='w')
connection_label=ttk.Label(settings_tab,textvariable=connection_status,font=('Segoe UI',13,'bold'),foreground='#64748b')
connection_label.pack(anchor='w',pady=5)
ttk.Label(settings_tab,textvariable=connection_detail,wraplength=530).pack(anchor='w')
load_api_settings()
def invalidate_connection(*args):
    connection_status.set('Not tested for current settings')
    connection_detail.set('')
    connection_label.configure(foreground='#64748b')
for variable in (base,model,api_key): variable.trace_add('write',invalidate_connection)

keys=keyboard.GlobalHotKeys({f'<{k}>':lambda fn=fn:(stop.set(),pause.clear(),commands.put((fn,None))) if fn is halt else commands.put((fn,None)) for k,fn in [('f6',record),('f7',finish),('f8',play),('f9',toggle),('f10',halt),('f12',close)]})
keys.start()
# Stop directly from the listener even while an approval dialog is open.

root.protocol('WM_DELETE_WINDOW',close)
root.after(50,poll)
root.mainloop()
