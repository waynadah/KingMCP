import tkinter as tk
import customtkinter as ctk
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
from tkinter import ttk, filedialog, messagebox
import threading, time, json, queue
from pynput import mouse, keyboard

root = ctk.CTk()
root.title('Desktop AI Agent v5.1 — Voice Edition')
root.geometry('1060x820')
root.minsize(860, 680)
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
rpm_var=tk.StringVar(value="8")
timeout_var=tk.StringVar(value="90")
request_status=tk.StringVar(value="API requests: not started")

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
        rpm_var.set(str(data.get('rpm',8)))
        timeout_var.set(str(data.get('timeout',90)))
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
        url, key, name, rpm, timeout = settings()
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
            'rpm': rpm, 'timeout': timeout, 'remember_key': bool(remember_key.get())}, indent=2), encoding='utf-8')
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
        if text.startswith('API requests:'): request_status.set(text)
        if text.startswith('Step '): ai_status.set('Analyzing screen...')
        elif text.startswith(('API cooldown:', 'Rate limit reached', 'API timeout or connection failure')): ai_status.set('Waiting for API cooldown...')
        elif text.startswith('Waiting for AI response'): ai_status.set('Waiting for AI response...')
        elif text.startswith('Verifying task'): ai_status.set('Checking task result...')
        elif text.startswith('Recovering from'): ai_status.set('Checking tool failure / replanning...')
        elif text.startswith('AI plan / update'): ai_status.set('Checking task plan...')
        elif text.startswith('Continuing:'): ai_status.set('Requesting next action...')
        elif text.startswith('AI response'): ai_status.set('Checking AI response...')
        elif text.split(' ',1)[0] in ('click','move','type_text','hotkey','press','scroll','list_windows','focus_window','excel_write','screen_size','draw_path','open_application','inspect_window','click_control','excel_read','excel_format'): ai_status.set('Executing / awaiting approval: '+text.split(' ',1)[0])
        log.insert('end', text+'\n'); log.see('end')
        if text.startswith(('AI plan / update:', 'AI response (no tool call):')):
            add_chat('Assistant', text.split(':',1)[1].strip())
    commands.put((update,None))

def settings():
    url=base.get().strip(); key=api_key.get().strip(); name=model.get().strip()
    if not url.startswith('https://') or not key or not name:
        raise ValueError('Enter an HTTPS base URL, API key and model ID.')
    rpm=int(rpm_var.get()); timeout=int(timeout_var.get())
    if not 1<=rpm<=10 or not 15<=timeout<=300: raise ValueError("RPM must be 1–10 and timeout 15–300 seconds.")
    return url,key,name,rpm,timeout

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
        commands.put((lambda: (ai_status.set(final_status),refresh_memory(),add_chat('Status',final_status), speak_result(final_status)),None))
        commands.put(('done',None))

def refresh_memory():
    try: memory_status.set(f'Memory: {len(task_memory.load())} saved entries')
    except Exception: memory_status.set('Memory unavailable — reset to repair')

def reset_tasks():
    if busy():
        messagebox.showinfo('Reset tasks','Stop the active task and wait until it finishes first.')
        return
    if not messagebox.askyesno('Reset all tasks','Clear saved task history, workflows, instructions and action log? API settings and recordings will remain.'): return
    try: task_memory.reset()
    except Exception as exc:
        messagebox.showerror('Reset failed',str(exc)); return
    log.delete('1.0','end'); task_box.delete('1.0','end')
    ai_status.set('Ready'); refresh_memory(); refresh_workflows(); reload_chat()

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
    if busy() or recording or voice_active or status.get()=='Starting recording...': return
    try:
        config=settings(); task=task_box.get('1.0','end').strip(); steps=int(limit.get())
        if not task or not 1<=steps<=100: raise ValueError('Enter a task and step limit (1–100).')
    except Exception as exc: messagebox.showerror('Settings',str(exc)); return
    if not messagebox.askyesno('Screen sharing','Screenshots of your primary display will be sent to your configured API provider. Close private windows first. Continue?'): return
    approval = bool(approve.get())
    if not approval and not messagebox.askyesno('Automatic mode', 'Desktop actions will execute without individual approval. Keep F10 ready to stop. Continue?'):
        return
    add_chat('You',task)
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


BG='#111827'; PANEL='#1f2937'; FG='#e5e7eb'; MUTED='#9ca3af'
style=ttk.Style(root)
style.theme_use('clam')
style.configure('.',background=BG,foreground=FG,fieldbackground=PANEL,font=('Segoe UI',10))
style.configure('TButton',padding=8,background=PANEL,foreground=FG)
style.map('TButton',background=[('active','#374151')])
style.configure('TFrame',background=BG)
style.configure('TLabel',background=BG,foreground=FG)
style.configure('TCheckbutton',background=BG,foreground=FG)
style.map('TCheckbutton',background=[('active',BG)])
style.configure('TEntry',fieldbackground=PANEL,foreground=FG,insertcolor=FG)
style.configure('TSpinbox',fieldbackground=PANEL,foreground=FG,insertcolor=FG)
style.configure('Horizontal.TProgressbar',background='#6366f1',troughcolor=PANEL)
root.option_add('*Text.background',PANEL)
root.option_add('*Text.foreground',FG)
root.option_add('*Text.insertBackground',FG)
root.option_add('*Listbox.background',PANEL)
root.option_add('*Listbox.foreground',FG)
root.option_add('*Listbox.selectBackground','#4338ca')
root.option_add('*Text.font','SegoeUI 10')
sidebar=ctk.CTkFrame(root,width=200,corner_radius=0,fg_color='#0b1220')
sidebar.pack(side='left',fill='y'); sidebar.pack_propagate(False)
ctk.CTkLabel(sidebar,text='Desktop AI',font=('Segoe UI',24,'bold')).pack(pady=(26,0))
ctk.CTkLabel(sidebar,text='v5.0  /  MCP Edition',text_color=MUTED).pack(pady=(0,24))
class SidebarPages(ttk.Frame):
    def __init__(self,parent):
        super().__init__(parent); self.pages={}; self.buttons={}; self.current=None
    def add(self,page,text):
        self.pages[page]=text
        button=ctk.CTkButton(sidebar,text=text,anchor='w',height=42,fg_color='transparent',hover_color='#263449',command=lambda:self.select(page))
        button.pack(fill='x',padx=12,pady=3); self.buttons[page]=button
    def select(self,page):
        if self.current is not None: self.current.pack_forget(); self.buttons[self.current].configure(fg_color='transparent')
        self.current=page; page.pack(fill='both',expand=True); self.buttons[page].configure(fg_color='#3730a3')
notebook=SidebarPages(root); notebook.pack(side='left',fill='both',expand=True,padx=12,pady=12)
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
ttk.Label(agent,textvariable=request_status).pack(anchor="w")
ttk.Label(agent,textvariable=ai_status,font=('Segoe UI',11,'bold'),foreground='#a5b4fc').pack(anchor='w')
activity=ttk.Progressbar(agent,mode='indeterminate')
activity.pack(fill='x',pady=(3,5))
def update_activity(*args):
    active=ai_status.get().startswith(('Starting','Analyzing','Waiting','Requesting','Checking','Executing'))
    if active: activity.start(15)
    else: activity.stop()
ai_status.trace_add('write',update_activity)
chat=ctk.CTkScrollableFrame(agent,fg_color='#0b1220',corner_radius=12)
chat.pack(fill='both',expand=True,pady=(8,10))
def add_chat(role,text):
    bubble=ctk.CTkFrame(chat,fg_color='#3730a3' if role=='You' else PANEL,corner_radius=10)
    bubble.pack(fill='x',padx=(48,8) if role=='You' else (8,48),pady=5)
    ctk.CTkLabel(bubble,text=role,font=('Segoe UI',11,'bold'),text_color='#c7d2fe').pack(anchor='w',padx=12,pady=(8,0))
    label=ctk.CTkLabel(bubble,text=str(text),justify='left',anchor='w',wraplength=450,font=('Segoe UI',13))
    label.pack(fill='x',padx=12,pady=(2,10))
    bubble.bind('<Configure>',lambda e:label.configure(wraplength=max(120,e.width-28)))
    children=chat.winfo_children()
    if len(children)>100: children[0].destroy()
    chat.after_idle(lambda:chat._parent_canvas.yview_moveto(1))
def reload_chat():
    for child in chat.winfo_children(): child.destroy()
    try:
        for row in task_memory.load():
            role={'user task':'You','AI response':'Assistant','AI plan / update':'Assistant','task outcome':'Status'}.get(row['kind'])
            if role: add_chat(role,row['text'])
    except Exception: add_chat('Status','Unable to load saved conversation.')
    if not chat.winfo_children(): add_chat('Assistant','Ready to help. Enter a desktop task below.')
reload_chat()
log_panel=ttk.Frame(agent)
log=tk.Text(log_panel,height=8,wrap='word'); log.pack(side='left',fill='both',expand=True)
scroll=ttk.Scrollbar(log_panel,command=log.yview); scroll.pack(side='right',fill='y'); log.configure(yscrollcommand=scroll.set)
def toggle_log():
    if log_panel.winfo_ismapped(): log_panel.pack_forget(); log_button.configure(text='Show action log')
    else: log_panel.pack(fill='x',before=composer,pady=4); log_button.configure(text='Hide action log')
log_button=ctk.CTkButton(agent,text='Show action log',command=toggle_log,fg_color=PANEL); log_button.pack(anchor='e',pady=3)
composer=ttk.Frame(agent); composer.pack(fill='x')
ttk.Label(composer,text='Task instructions').pack(anchor='w')
task_box=tk.Text(composer,height=3,wrap='word'); task_box.pack(fill='x',pady=4)
def send_shortcut(event):
    start_ai(); return 'break'
task_box.bind('<Control-Return>',send_shortcut)

# Voice input is opt-in and never automatically submits a task.
voice_active = False
voice_language = tk.StringVar(value='en-US')
voice_output = tk.BooleanVar(value=False)
voice_status = tk.StringVar(value='Voice ready')

def speak_result(text):
    if not voice_output.get() or closing: return
    def run():
        try:
            from voice_tools import speak
            speak(text)
        except Exception as exc:
            commands.put((lambda msg=str(exc): voice_status.set('Speech error: '+msg),None))
    threading.Thread(target=run,daemon=True).start()

def voice_input():
    global voice_active
    if voice_active or busy() or recording: return
    if not messagebox.askyesno('Microphone privacy', 'Record one short instruction and send its audio to Google speech recognition for transcription? Review the text before sending the task.',parent=root): return
    voice_active=True
    language=voice_language.get()
    mic_button.configure(state='disabled')
    voice_status.set('Listening / transcribing...')
    def finish_voice(text,error):
        global voice_active
        voice_active=False
        mic_button.configure(state='normal')
        if error:
            voice_status.set('Voice error: '+error)
        elif busy() or recording:
            voice_status.set('Task active; transcript not inserted.')
        else:
            task_box.insert('end', (' ' if task_box.get('1.0','end').strip() else '')+text)
            voice_status.set('Review transcript, then Send task.')
    def run():
        try:
            from voice_tools import listen
            text=listen(language)
            commands.put((lambda:finish_voice(text,None),None))
        except Exception as exc:
            commands.put((lambda msg=str(exc) or type(exc).__name__:finish_voice('',msg),None))
    threading.Thread(target=run,daemon=True).start()

voice_line=ttk.Frame(composer); voice_line.pack(fill='x',pady=3)
mic_button=ctk.CTkButton(voice_line,text='Speak task',width=110,command=voice_input)
mic_button.pack(side='left',padx=2)
ttk.Combobox(voice_line,textvariable=voice_language,values=('en-US','en-PH','fil-PH'),state='readonly',width=9).pack(side='left',padx=4)
ttk.Checkbutton(voice_line,text='Speak task status',variable=voice_output).pack(side='left',padx=4)
ttk.Label(composer,textvariable=voice_status,wraplength=600).pack(anchor='w')

line=ttk.Frame(composer); line.pack(fill='x')
ttk.Checkbutton(line,text='Approve each action',variable=approve).pack(side='left')
ttk.Label(line,text='Max steps').pack(side='left',padx=(12,2)); ttk.Spinbox(line,from_=1,to=100,textvariable=limit,width=5).pack(side='left')
line=ttk.Frame(composer); line.pack(fill='x',pady=6)
for label,fn in [('Send task',start_ai),('Pause / Resume F9',toggle),('Stop F10',halt)]:
    ctk.CTkButton(line,text=label,command=fn,fg_color='#4338ca' if fn is start_ai else PANEL).pack(side='left',expand=True,fill='x',padx=2)
memory_status=tk.StringVar()
memline=ttk.Frame(composer); memline.pack(fill='x')
ttk.Label(memline,textvariable=memory_status).pack(side='left')
ttk.Button(memline,text='Reset All Tasks',command=reset_tasks).pack(side='right')
ttk.Button(memline,text='Continue Task',command=continue_task).pack(side='right',padx=4)
refresh_memory()
tools_tab=ttk.Frame(notebook,padding=16); notebook.add(tools_tab,text='MCP Tools')
ttk.Label(tools_tab,text='Local MCP server (stdio)',font=('Segoe UI',18,'bold')).pack(anchor='w')
ttk.Label(tools_tab,text='Starts automatically for each AI task. No network listener.',wraplength=530).pack(anchor='w',pady=10)
tool_list=tk.Listbox(tools_tab,exportselection=False); tool_list.pack(fill='both',expand=True)
for name in ['screenshot','screen_size','click','move','type_text','hotkey','press','scroll','list_windows','focus_window','excel_write','draw_path','open_application','inspect_window','click_control','excel_read','excel_format','browser_open','file_list','file_read','file_write','file_search','word_create','word_append','excel_add_sheet','excel_formula','excel_save_copy']:
    tool_list.insert('end',name)
workflow_tab=ttk.Frame(notebook,padding=16); notebook.add(workflow_tab,text='Workflows')
ttk.Label(workflow_tab,text='Reusable workflows',font=('Segoe UI',18,'bold')).pack(anchor='w')
ttk.Label(workflow_tab,text='Save instructions and run them again with fresh desktop analysis. No blind replay.',wraplength=530).pack(anchor='w',pady=8)
workflow_list=tk.Listbox(workflow_tab,height=8,exportselection=False); workflow_list.pack(fill='both',expand=True)
workflow_name=tk.StringVar()
ttk.Entry(workflow_tab,textvariable=workflow_name).pack(fill='x',pady=5)
def refresh_workflows():
    workflow_list.delete(0,'end')
    try:
        for name,prompt in task_memory.workflows(): workflow_list.insert('end',name)
    except Exception as exc: messagebox.showerror('Workflows',str(exc))
def save_workflow():
    if busy(): return
    try:
        prompt=task_memory.last_successful_prompt()
        if not prompt: raise ValueError('Finish a task successfully before saving it.')
        name=workflow_name.get().strip()
        if name in dict(task_memory.workflows()) and not messagebox.askyesno('Replace workflow','Replace this saved workflow?'): return
        task_memory.save_workflow(name,prompt); refresh_workflows()
    except Exception as exc: messagebox.showerror('Workflows',str(exc))
def selected_workflow():
    selection=workflow_list.curselection()
    return workflow_list.get(selection[0]) if selection else None
def load_workflow():
    if busy(): return
    name=selected_workflow()
    if name:
        prompt=dict(task_memory.workflows()).get(name)
        if prompt:
            task_box.delete('1.0','end'); task_box.insert('1.0',prompt); notebook.select(agent)
def remove_workflow():
    if busy(): return
    name=selected_workflow()
    if name and messagebox.askyesno('Delete workflow','Delete '+name+'?'):
        task_memory.delete_workflow(name); refresh_workflows()
for label,fn in [('Save last successful task',save_workflow),('Load selected into Agent',load_workflow),('Delete selected',remove_workflow)]:
    ttk.Button(workflow_tab,text=label,command=fn).pack(fill='x',pady=3)
refresh_workflows()
settings_tab=ttk.Frame(notebook,padding=16); notebook.add(settings_tab,text='API Settings')
settings_entries=[]
for label,var,secret in [('Base URL',base,False),('API Key',api_key,True),('Model ID',model,False),('Requests per minute (1–10)',rpm_var,False),('Timeout seconds (15–300)',timeout_var,False)]:
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

memory_tab=ttk.Frame(notebook,padding=16); notebook.add(memory_tab,text='Memory')
ttk.Label(memory_tab,text='Saved task history',font=('Segoe UI',20,'bold')).pack(anchor='w')
ttk.Label(memory_tab,textvariable=memory_status).pack(anchor='w',pady=6)
memory_view=tk.Text(memory_tab,wrap='word'); memory_view.pack(fill='both',expand=True)
def show_memory():
    memory_view.configure(state='normal'); memory_view.delete('1.0','end')
    try:
        for row in task_memory.load(): memory_view.insert('end',row['kind']+'\n'+row['text']+'\n\n')
    except Exception as exc: memory_view.insert('end',str(exc))
    memory_view.configure(state='disabled')
ttk.Button(memory_tab,text='Refresh history',command=show_memory).pack(fill='x',pady=4)
ttk.Button(memory_tab,text='Reset all tasks and workflows',command=lambda:(reset_tasks(),show_memory())).pack(fill='x')
show_memory()
notebook.select(agent)
ctk.CTkLabel(sidebar,text='F10  •  Emergency stop',text_color=MUTED).pack(side='bottom',pady=16)
keys=keyboard.GlobalHotKeys({f'<{k}>':lambda fn=fn:(stop.set(),pause.clear(),commands.put((fn,None))) if fn is halt else commands.put((fn,None)) for k,fn in [('f6',record),('f7',finish),('f8',play),('f9',toggle),('f10',halt),('f12',close)]})
keys.start()
# Stop directly from the listener even while an approval dialog is open.

root.protocol('WM_DELETE_WINDOW',close)
root.after(50,poll)
root.mainloop()
