"""Bounded local task history. No screenshots or API credentials are stored."""
import json, os
from pathlib import Path
PATH = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'DesktopAI' / 'task_history.json'
MAX = 100

def load():
    if not PATH.exists(): return []
    data = json.loads(PATH.read_text(encoding='utf-8'))
    if not isinstance(data,list): raise ValueError('Invalid task history')
    return data[-MAX:]

def append(kind, text):
    data = load()
    data.append({'kind':kind,'text':str(text)[:3000]})
    PATH.parent.mkdir(parents=True,exist_ok=True)
    tmp = PATH.with_suffix('.tmp')
    tmp.write_text(json.dumps(data[-MAX:],ensure_ascii=False),encoding='utf-8')
    os.replace(tmp,PATH)

def reset():
    PATH.unlink(missing_ok=True)
    PATH.with_suffix('.tmp').unlink(missing_ok=True)

def context():
    return json.dumps(load(),ensure_ascii=False)[-24000:]
