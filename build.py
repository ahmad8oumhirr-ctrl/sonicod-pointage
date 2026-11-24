import pathlib, re, sys
src = pathlib.Path('src')
def lit(t): return '`' + t.replace('\\','\\\\').replace('`','\\`').replace('${','\\${') + '`'
scan = src.joinpath('scan.html').read_text(); admin = src.joinpath('admin.html').read_text()
out = src.joinpath('server.js').read_text()
out += '\n/* ====================== HTML (ma tbddel walo hna) ====================== */\n'
out += 'const SCAN_HTML = ' + lit(scan) + ';\n\nconst ADMIN_HTML = ' + lit(admin) + ';\n'
pathlib.Path('Code.gs').write_text(out)

# ---- version site statique (GitHub Pages) ----
api = pathlib.Path('api_url.txt').read_text().strip() if pathlib.Path('api_url.txt').exists() else 'COLLER_LIEN_EXEC_ICI'
FETCH = """function call(fn){
  const args = Array.prototype.slice.call(arguments, 1);
  return fetch(API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ fn: fn, args: args }) })
    .then(r => r.json()).then(j => { if (j.err) throw new Error(j.err); return j.r; });
}"""
OLD = re.compile(r"function call\(fn\)\{\n.*?\n\}", re.S)
def static(h, extra):
    h = OLD.sub(lambda m: FETCH, h, count=1)
    h = h.replace('<base target="_top">', '<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">\n<title>Sonicod Pointage</title>')
    h = h.replace('<script>\nconst', '<script>\n' + extra + 'const', 1)
    return h.replace('__AGENCE__', 'Agadir Tilila')
d = pathlib.Path('.')
SCAN_API = "const API_ID = (new URLSearchParams(location.search).get('s') || '').replace(/[^A-Za-z0-9_-]/g, '');\nconst API_URL = 'https://script.google.com/macros/s/' + API_ID + '/exec';\n"
s = static(scan, SCAN_API).replace("const CODE = '__AGENCE_CODE__';", "const CODE = (new URLSearchParams(location.search).get('a') || '').replace(/[^A-Za-z0-9-]/g, '');")
d.joinpath('index.html').write_text(s)
ADMIN_API = """const STATIC_SITE = new URL('./', location.href).href;
let API_ID = ''; try { API_ID = localStorage.getItem('sonicod_pt_api') || ''; } catch(e){}
let API_URL = 'https://script.google.com/macros/s/' + API_ID + '/exec';
"""
a = static(admin, ADMIN_API).replace("const APP_URL = '__APP_URL__';", "const APP_URL = '';")
a = a.replace('<title>Sonicod Pointage</title>', '<title>Sonicod Pointage — Admin</title>')
d.joinpath('admin.html').write_text(a)
