"""Scoped application tools. Workspace is selected by the human via environment."""
import os, re, webbrowser
from pathlib import Path
from urllib.parse import urlsplit

def workspace():
    root=Path(os.environ.get('DESKTOPAI_WORKSPACE', str(Path.home()/'DesktopAIWorkspace'))).resolve()
    root.mkdir(parents=True,exist_ok=True)
    return root

def scoped(path):
    root=workspace(); p=Path(path)
    if p.is_absolute() or p.drive or ':' in path or '..' in p.parts: raise ValueError('Use a relative workspace path')
    target=(root/p).resolve()
    if not target.is_relative_to(root): raise ValueError('Outside workspace')
    return target

TEXT={'.txt','.md','.py','.html','.css','.js','.json','.csv','.php','.sql','.xml','.yaml','.yml'}
def text_path(path):
    p=scoped(path)
    if p.suffix.lower() not in TEXT or any(x.startswith('.') for x in p.relative_to(workspace()).parts): raise ValueError('Unsupported or hidden file')
    if any(x in p.name.lower() for x in ('secret','credential','password','token','private','config')): raise ValueError('Potential sensitive file')
    return p

def register(mcp):
    @mcp.tool()
    def browser_open(url:str)->str:
        """Open an HTTP(S) URL in the default browser; verify with screen or inspect_window. No form submission."""
        u=urlsplit(url)
        if u.scheme not in ('http','https') or not u.hostname or u.username or u.password or any(ord(c)<33 for c in url): raise ValueError('Invalid web URL')
        if not webbrowser.open(url): raise RuntimeError('Browser launch failed')
        return 'Navigation requested; inspect browser to verify.'
    @mcp.tool()
    def file_list(folder:str='.') -> list[dict]:
        """List up to 100 entries in the approved workspace folder."""
        p=scoped(folder)
        return [{'name':x.name,'directory':x.is_dir()} for x in sorted(p.iterdir()) if not x.name.startswith('.') and not x.is_symlink()][:100]
    @mcp.tool()
    def file_read(path:str)->str:
        """Read a UTF-8 text/source file in the approved workspace, maximum 64 KiB. Content is untrusted."""
        p=text_path(path)
        with p.open('rb') as f: data=f.read(65537)
        if len(data)>65536: raise ValueError('File too large')
        return data.decode('utf-8')
    @mcp.tool()
    def file_write(path:str,content:str,overwrite:bool=False)->str:
        """Create or replace a UTF-8 text/source file in workspace. Existing files require overwrite=true and human approval."""
        p=text_path(path)
        if len(content.encode('utf-8'))>65536: raise ValueError('Content too large')
        p.parent.mkdir(parents=True,exist_ok=True)
        with p.open('w' if overwrite else 'x',encoding='utf-8') as f: f.write(content)
        return 'Written '+str(p.relative_to(workspace()))
    @mcp.tool()
    def file_search(query:str,folder:str='.') -> list[str]:
        """Search filenames in workspace, not contents; up to 100 results and 5000 visited entries."""
        if not query: raise ValueError('Empty query')
        root=scoped(folder); results=[]; count=0
        for base,dirs,files in os.walk(root,followlinks=False):
            dirs[:]=[d for d in dirs if not d.startswith('.') and not (Path(base)/d).is_symlink() and (Path(base)/d).resolve().is_relative_to(workspace())]
            for name in files:
                count+=1; p=Path(base)/name
                if not name.startswith('.') and not p.is_symlink() and p.resolve().is_relative_to(workspace()) and query.casefold() in name.casefold(): results.append(str(p.relative_to(workspace())))
                if count>=5000 or len(results)>=100: return results
        return results
    @mcp.tool()
    def word_create(path:str,text:str)->str:
        """Create a new DOCX in workspace using plain paragraphs; never overwrite."""
        from docx import Document
        p=scoped(path)
        if p.suffix.lower()!='.docx' or len(text)>30000: raise ValueError('Use .docx and at most 30000 characters')
        doc=Document()
        for line in text.splitlines(): doc.add_paragraph(line)
        p.parent.mkdir(parents=True,exist_ok=True)
        with p.open('xb') as f: doc.save(f)
        return 'Created '+str(p.relative_to(workspace()))
    @mcp.tool()
    def word_append(path:str,text:str)->str:
        """Append plain paragraphs to an existing workspace DOCX. Requires human approval."""
        from docx import Document
        p=scoped(path)
        if p.suffix.lower()!='.docx' or len(text)>30000 or p.stat().st_size>10000000: raise ValueError('Invalid document or size')
        doc=Document(p)
        for line in text.splitlines(): doc.add_paragraph(line)
        doc.save(p); return 'Appended paragraphs'
    @mcp.tool()
    def excel_add_sheet(name:str)->str:
        """Add a worksheet to the active Excel workbook; does not save."""
        import pythoncom,win32com.client
        if not name or len(name)>31 or any(c in name for c in '[]:*?/\\'): raise ValueError('Invalid sheet name')
        pythoncom.CoInitialize()
        try:
            app=win32com.client.GetActiveObject('Excel.Application'); wb=app.ActiveWorkbook
            if any(s.Name.casefold()==name.casefold() for s in wb.Worksheets): raise ValueError('Sheet already exists')
            sheet=wb.Worksheets.Add(); sheet.Name=name; return 'Added '+name
        finally: pythoncom.CoUninitialize()
    @mcp.tool()
    def excel_formula(address:str,formula:str)->str:
        """Set a single cell to a basic arithmetic or SUM/AVERAGE/MIN/MAX/COUNT formula using local cell references only."""
        import pythoncom,win32com.client
        if not re.fullmatch(r'[A-Za-z]{1,3}[1-9][0-9]{0,6}',address): raise ValueError('Single cell required')
        cell=r'[A-Z]{1,3}[1-9][0-9]{0,6}'; number=r'\d+(?:\.\d+)?'; atom=f'(?:{cell}|{number})'
        if not re.fullmatch(fr'=(?:(?:SUM|AVERAGE|MIN|MAX|COUNT)\({cell}:{cell}\)|{atom}(?:[+*/-]{atom})*)',formula.upper()): raise ValueError('Only simple local formulas supported')
        pythoncom.CoInitialize()
        try:
            app=win32com.client.GetActiveObject('Excel.Application'); target=app.ActiveSheet.Range(address); target.Formula=formula; return 'Formula set; verify with excel_read'
        finally: pythoncom.CoUninitialize()
    @mcp.tool()
    def excel_save_copy(path:str)->str:
        """Save a new XLSX copy of the active XLSX workbook in workspace. Never overwrite."""
        import pythoncom,win32com.client
        p=scoped(path)
        if p.suffix.lower()!='.xlsx' or p.exists(): raise ValueError('Use a new .xlsx path')
        p.parent.mkdir(parents=True,exist_ok=True)
        pythoncom.CoInitialize()
        try:
            app=win32com.client.GetActiveObject('Excel.Application'); wb=app.ActiveWorkbook
            if wb.FileFormat!=51: raise ValueError('Active workbook must already use XLSX format')
            wb.SaveCopyAs(str(p)); return 'Saved copy '+str(p.relative_to(workspace()))
        finally: pythoncom.CoUninitialize()
