"""SQLite task history and reusable prompt workflows; no screenshot storage."""
import json, os, sqlite3
from pathlib import Path
from contextlib import contextmanager
PATH = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'DesktopAI' / 'task_history.json'
DB = PATH.with_name('memory.db')
MAX = 100
@contextmanager
def connection():
    DB.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB, timeout=10)
    try:
        with con:
            con.execute('PRAGMA secure_delete=ON')
            con.execute('CREATE TABLE IF NOT EXISTS history (id INTEGER PRIMARY KEY, kind TEXT, text TEXT)')
            con.execute('CREATE TABLE IF NOT EXISTS workflows (name TEXT PRIMARY KEY, prompt TEXT NOT NULL)')
            con.execute('CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY)')
            if not con.execute("SELECT 1 FROM meta WHERE key='migration'").fetchone():
                if PATH.exists():
                    rows=json.loads(PATH.read_text(encoding='utf-8'))
                    if not isinstance(rows,list): raise ValueError('Invalid legacy history')
                    con.executemany('INSERT INTO history(kind,text) VALUES (?,?)', [(str(r['kind']),str(r['text'])[:3000]) for r in rows[-MAX:]])
                con.execute("INSERT INTO meta VALUES ('migration')")
            yield con
    finally: con.close()
def load():
    with connection() as c:
        return [{'kind':k,'text':t} for k,t in reversed(c.execute('SELECT kind,text FROM history ORDER BY id DESC LIMIT ?', (MAX,)).fetchall())]
def append(kind,text):
    with connection() as c:
        c.execute('INSERT INTO history(kind,text) VALUES (?,?)',(kind,str(text)[:3000]))
        c.execute('DELETE FROM history WHERE id NOT IN (SELECT id FROM history ORDER BY id DESC LIMIT ?)',(MAX,))
def reset():
    with connection() as c:
        c.execute('DELETE FROM history'); c.execute('DELETE FROM workflows')
    PATH.unlink(missing_ok=True); PATH.with_suffix('.tmp').unlink(missing_ok=True)
def context():
    rows=load(); selected=[]
    for row in reversed(rows):
        if len(json.dumps([row]+selected,ensure_ascii=False))>24000: break
        selected.insert(0,row)
    return json.dumps(selected,ensure_ascii=False)
def save_workflow(name,prompt):
    if not name.strip() or not prompt.strip(): raise ValueError('Name and instructions required')
    with connection() as c: c.execute('INSERT OR REPLACE INTO workflows VALUES (?,?)',(name.strip(),prompt.strip()))
def workflows():
    with connection() as c: return c.execute('SELECT name,prompt FROM workflows ORDER BY name COLLATE NOCASE').fetchall()
def delete_workflow(name):
    with connection() as c: c.execute('DELETE FROM workflows WHERE name=?',(name,))
def last_successful_prompt():
    rows=load()
    if not rows or rows[-1]['kind']!='task outcome' or not rows[-1]['text'].startswith('Finished'): return None
    return next((r['text'] for r in reversed(rows) if r['kind']=='user task'),None)
