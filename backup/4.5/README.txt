Desktop AI Agent v4 — MCP Edition (Windows prototype)

Extract the ZIP, open CMD in the extracted DesktopAI folder:
cd /d E:\Phyton\auto_clicker
.venv\Scripts\activate
  py -m venv .venv
  .venv\Scripts\python -m pip install -r requirements.txt
  .venv\Scripts\python main.py

API Settings: enter your key, test API and native tool calling. Then use AI Agent.
The configured provider must support both native function calls and image input.
There is no JSON fallback in this version. API keys are not saved.
Screenshots and tool results are sent to the configured provider. Close private windows.
Manual approval is required. F10 stops future actions; an in-flight API or tool call may finish first. F9 pauses/resumes, F12 exits. PyAutoGUI corner fail-safe applies to mouse/keyboard tools, not Excel COM.
Excel tool writes literal text to the active worksheet and can overwrite cells. It does not save. Excel must be installed and running, with an editable workbook.
MCP uses a local stdio subprocess, not a network endpoint. External trusted MCP clients can launch desktop_server.py directly, but must provide their own approval and safety controls. Do not expose it remotely.
Retains recorder, JSON save/load and speed controls. No persistent AI memory, browser DOM tools, or UI accessibility automation yet.
