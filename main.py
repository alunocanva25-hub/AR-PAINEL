from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
import sqlite3, shutil, re, json
import urllib.request
import requests
from datetime import datetime

BASE = Path(__file__).parent
UPLOADS = BASE / 'uploads'
DB = BASE / 'database' / 'ar_panel.db'
DATABASE_DIR = BASE / 'database'
UPLOADS.mkdir(exist_ok=True)
DATABASE_DIR.mkdir(exist_ok=True)
XLSX_CACHE = DATABASE_DIR / 'base_import.xlsx'
HEADER_ROW = 5
START_COL = 2  # Coluna B

app = FastAPI(title='DSYSTEM AR PANEL', version='1.0.0.8')
app.mount('/static', StaticFiles(directory=str(BASE/'static')), name='static')

def conn():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    return c

def init_db():
    with conn() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS ars(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            instalacao TEXT DEFAULT '',
            medidor TEXT DEFAULT '',
            nome_cliente TEXT DEFAULT '',
            filename TEXT NOT NULL,
            original_filename TEXT DEFAULT '',
            status TEXT DEFAULT 'Pendente',
            operador_usuario TEXT DEFAULT '',
            operador_nome TEXT DEFAULT '',
            operador_perfil TEXT DEFAULT ''
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS base_xlsx(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            aba TEXT DEFAULT '',
            instalacao TEXT DEFAULT '',
            medidor TEXT DEFAULT '',
            nome_cliente TEXT DEFAULT ''
        )""")
        cols_ars = [r['name'] for r in c.execute('PRAGMA table_info(ars)').fetchall()]
        if 'nome_cliente' not in cols_ars:
            c.execute("ALTER TABLE ars ADD COLUMN nome_cliente TEXT DEFAULT ''")
        if 'api_id' not in cols_ars:
            c.execute("ALTER TABLE ars ADD COLUMN api_id TEXT DEFAULT ''")
        if 'api_url' not in cols_ars:
            c.execute("ALTER TABLE ars ADD COLUMN api_url TEXT DEFAULT ''")
        if 'operador_usuario' not in cols_ars:
            c.execute("ALTER TABLE ars ADD COLUMN operador_usuario TEXT DEFAULT ''")
        if 'operador_nome' not in cols_ars:
            c.execute("ALTER TABLE ars ADD COLUMN operador_nome TEXT DEFAULT ''")
        if 'operador_perfil' not in cols_ars:
            c.execute("ALTER TABLE ars ADD COLUMN operador_perfil TEXT DEFAULT ''")
        cols_base = [r['name'] for r in c.execute('PRAGMA table_info(base_xlsx)').fetchall()]
        if 'aba' not in cols_base:
            c.execute("ALTER TABLE base_xlsx ADD COLUMN aba TEXT DEFAULT ''")
        if 'nome_cliente' not in cols_base:
            c.execute("ALTER TABLE base_xlsx ADD COLUMN nome_cliente TEXT DEFAULT ''")
        c.commit()
init_db()

def safe(s):
    s = re.sub(r'[^A-Za-z0-9_.-]+','_',str(s or ''))
    return s.strip('_')

def next_pag(ext='pdf'):
    files = list(UPLOADS.glob('AR_CARTACONVITE_PAG.*'))
    nums=[]
    for f in files:
        m=re.search(r'PAG\.(\d+)',f.name)
        if m: nums.append(int(m.group(1)))
    return f'AR_CARTACONVITE_PAG.{(max(nums)+1 if nums else 1):02d}.{ext}'

def build_name(inst, md, ext='pdf'):
    inst = safe(inst)
    md = safe(md)
    if inst or md:
        return f'AR_CARTACONVITE_INST_{inst}_MD_{md}.{ext}'
    return next_pag(ext)

def extract_from_filename(name):
    txt = str(name or '')
    m = re.search(r'INST_([^_\.]+)_MD_([^_\.]+)', txt, re.I)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m = re.search(r'INST_([^_\.]+)', txt, re.I)
    inst = m.group(1).strip() if m else ''
    m = re.search(r'MD_([^_\.]+)', txt, re.I)
    md = m.group(1).strip() if m else ''
    return inst, md

def norm_header(v):
    s = str(v or '').strip().upper()
    mapa = {'Á':'A','À':'A','Â':'A','Ã':'A','É':'E','Ê':'E','Í':'I','Ó':'O','Õ':'O','Ô':'O','Ú':'U','Ç':'C'}
    for a,b in mapa.items():
        s=s.replace(a,b)
    return re.sub(r'\s+', ' ', s)

def cell_str(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def load_xlsx_book():
    try:
        from openpyxl import load_workbook
    except Exception:
        raise HTTPException(500, 'openpyxl não instalado')
    if not XLSX_CACHE.exists():
        raise HTTPException(404, 'Nenhuma base XLSX importada ainda')
    return load_workbook(XLSX_CACHE, data_only=True)

def auto_col(headers, names):
    for target in names:
        for i,h in enumerate(headers):
            if h == target or target in h:
                return i+1
    return None

@app.get('/', response_class=HTMLResponse)
def home():
    return FileResponse(BASE/'static'/'index.html')

@app.get('/api/ars')
def list_ars(instalacao: str='', medidor: str='', nome_cliente: str='', status: str='', operador: str='', data_ini: str='', data_fim: str=''):
    q='SELECT * FROM ars WHERE 1=1'; p=[]
    if instalacao:
        q+=' AND instalacao LIKE ?'; p.append(f'%{instalacao}%')
    if medidor:
        q+=' AND medidor LIKE ?'; p.append(f'%{medidor}%')
    if nome_cliente:
        q+=' AND nome_cliente LIKE ?'; p.append(f'%{nome_cliente}%')
    if status and status != 'Todos':
        q+=' AND status=?'; p.append(status)
    if operador:
        q+=' AND (operador_nome LIKE ? OR operador_usuario LIKE ?)'
        p.append(f'%{operador}%')
        p.append(f'%{operador}%')
    if data_ini:
        q+=' AND created_at >= ?'
        p.append(data_ini)
    if data_fim:
        q+=' AND created_at <= ?'
        p.append(data_fim + ' 23:59:59')
    q+=' ORDER BY id DESC'
    with conn() as c:
        return [dict(r) for r in c.execute(q,p).fetchall()]

@app.post('/api/upload')
def upload(file: UploadFile=File(...), instalacao: str=Form(''), medidor: str=Form(''), nome_cliente: str=Form(''), operador_usuario: str=Form(''), operador_nome: str=Form(''), operador_perfil: str=Form('')):
    ext = (Path(file.filename).suffix or '.pdf').lower().replace('.','')
    if ext not in ['pdf','jpg','jpeg','png','webp']:
        raise HTTPException(400, 'Tipo de arquivo não permitido')
    if not instalacao and not medidor:
        instalacao, medidor = extract_from_filename(file.filename)
    filename = build_name(instalacao, medidor, ext)
    dest = UPLOADS / filename
    base = dest.stem; i=2
    while dest.exists():
        dest = UPLOADS / f'{base}_{i}.{ext}'; i+=1
    with dest.open('wb') as out:
        shutil.copyfileobj(file.file, out)
    with conn() as c:
        c.execute('INSERT INTO ars(created_at, instalacao, medidor, nome_cliente, filename, original_filename, operador_usuario, operador_nome, operador_perfil) VALUES(?,?,?,?,?,?,?,?,?)',
                  (datetime.now().strftime('%d/%m/%Y %H:%M:%S'), instalacao, medidor, nome_cliente, dest.name, file.filename, operador_usuario, operador_nome, operador_perfil))
        c.commit()
    return {'ok': True, 'filename': dest.name}

@app.get('/api/download/{id}')
def download(id:int):
    with conn() as c:
        r=c.execute('SELECT * FROM ars WHERE id=?',(id,)).fetchone()
    if not r: raise HTTPException(404,'Não encontrado')
    path=UPLOADS/r['filename']
    if not path.exists(): raise HTTPException(404,'Arquivo não encontrado')
    return FileResponse(path, filename=r['filename'])

@app.get('/api/view/{id}')
def view(id:int):
    with conn() as c:
        r=c.execute('SELECT * FROM ars WHERE id=?',(id,)).fetchone()
    if not r: raise HTTPException(404,'Não encontrado')
    path=UPLOADS/r['filename']
    if not path.exists(): raise HTTPException(404,'Arquivo não encontrado')
    return FileResponse(path)

@app.post('/api/status/{id}')
def status(id:int, status: str=Form(...)):
    with conn() as c:
        c.execute('UPDATE ars SET status=? WHERE id=?',(status,id)); c.commit()
    return {'ok':True}

@app.post('/api/rename/{id}')
def rename(id:int, instalacao: str=Form(''), medidor: str=Form(''), nome_cliente: str=Form('')):
    with conn() as c:
        r=c.execute('SELECT * FROM ars WHERE id=?',(id,)).fetchone()
        if not r: raise HTTPException(404,'Não encontrado')
        old=UPLOADS/r['filename']
        ext=old.suffix.lower().replace('.','') or 'pdf'
        newname=build_name(instalacao, medidor, ext)
        new=UPLOADS/newname
        base=new.stem; i=2
        while new.exists() and new.name != old.name:
            new=UPLOADS/f'{base}_{i}.{ext}'; i+=1
        if old.exists() and new.name != old.name:
            old.rename(new)
        c.execute('UPDATE ars SET instalacao=?, medidor=?, nome_cliente=?, filename=? WHERE id=?',(instalacao, medidor, nome_cliente, new.name, id)); c.commit()
    return {'ok':True, 'filename': new.name}

@app.delete('/api/ars/{id}')
def delete(id:int):
    remote_ok = False
    remote_msg = ''
    with conn() as c:
        r=c.execute('SELECT * FROM ars WHERE id=?',(id,)).fetchone()
        if r:
            api_id = r['api_id'] if 'api_id' in r.keys() else ''
            api_url = r['api_url'] if 'api_url' in r.keys() else ''
            remote_ok, remote_msg = api_delete_remote(api_url, api_id)

            p=UPLOADS/r['filename']
            if p.exists(): p.unlink()
            c.execute('DELETE FROM ars WHERE id=?',(id,)); c.commit()
    return {'ok':True, 'api_deleted': remote_ok, 'api_message': remote_msg}


def api_json_get(url: str):
    req = urllib.request.Request(url, headers={'User-Agent': 'DSYSTEM-AR-PAINEL'})
    with urllib.request.urlopen(req, timeout=25) as resp:
        return json.loads(resp.read().decode('utf-8'))

def api_bytes_get(url: str):
    req = urllib.request.Request(url, headers={'User-Agent': 'DSYSTEM-AR-PAINEL'})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read()

def api_delete_remote(api_url: str, api_id: str):
    if not api_url or not api_id:
        return False, 'Sem vínculo com API'
    try:
        req = urllib.request.Request(
            api_url.rstrip('/') + f'/api/ars/{api_id}',
            method='DELETE',
            headers={'User-Agent': 'DSYSTEM-AR-PAINEL'}
        )
        with urllib.request.urlopen(req, timeout=20) as resp:
            resp.read()
        return True, 'Excluído da API'
    except Exception as e:
        return False, str(e)

@app.post('/api/receber-api')
def receber_api(api_url: str = Form(...)):
    api_url = (api_url or '').strip().rstrip('/')
    if not api_url:
        raise HTTPException(400, 'Informe a URL da API')

    try:
        registros = api_json_get(api_url + '/api/ars')
    except Exception as e:
        raise HTTPException(500, f'Falha ao consultar API: {e}')

    if not isinstance(registros, list):
        raise HTTPException(500, 'Resposta inválida da API')

    importados = 0
    ignorados = 0
    erros = []

    with conn() as c:
        for item in registros:
            try:
                api_id = str(item.get('id', '')).strip()
                if api_id:
                    ja = c.execute(
                        'SELECT id FROM ars WHERE api_id=? AND api_url=? LIMIT 1',
                        (api_id, api_url)
                    ).fetchone()
                    if ja:
                        ignorados += 1
                        continue

                file_name = item.get('file_name') or item.get('filename') or f'AR_API_{api_id or importados + 1}.pdf'
                ext = (Path(file_name).suffix or '.pdf').lower().replace('.', '')
                if ext not in ['pdf', 'jpg', 'jpeg', 'png', 'webp']:
                    ext = 'pdf'

                instalacao = str(item.get('instalacao') or '').strip()
                medidor = str(item.get('medidor') or '').strip()
                nome_cliente = str(item.get('nome_cliente') or '').strip()
                status = str(item.get('status') or 'Pendente').strip() or 'Pendente'
                operador_usuario = str(item.get('operador_usuario') or '').strip()
                operador_nome = str(item.get('operador_nome') or '').strip()
                operador_perfil = str(item.get('operador_perfil') or '').strip()

                final_name = build_name(instalacao, medidor, ext)
                dest = UPLOADS / final_name
                base_name = dest.stem
                n = 2
                while dest.exists():
                    dest = UPLOADS / f'{base_name}_{n}.{ext}'
                    n += 1

                try:
                    content = api_bytes_get(api_url + f'/api/ars/{api_id}/download')
                except Exception:
                    content = api_bytes_get(api_url + f'/api/download/{api_id}')

                dest.write_bytes(content)

                created_at = item.get('created_at') or datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                c.execute("""
                    INSERT INTO ars(
                        created_at, instalacao, medidor, nome_cliente,
                        filename, original_filename, status, api_id, api_url,
                        operador_usuario, operador_nome, operador_perfil
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    created_at, instalacao, medidor, nome_cliente,
                    dest.name, file_name, status, api_id, api_url,
                    operador_usuario, operador_nome, operador_perfil
                ))
                importados += 1
            except Exception as e:
                erros.append(f"{item.get('id', '?')}: {e}")

        c.commit()

    return {
        'ok': True,
        'importados': importados,
        'ignorados': ignorados,
        'erros': erros[:20],
        'total_api': len(registros)
    }


def panel_default_api_url():
    return os.environ.get('DSYSTEM_AR_API_URL', 'https://dsystem-ar-api.onrender.com').strip().rstrip('/')

def normalize_panel_api_url(api_url: str = ''):
    u = (api_url or '').strip().rstrip('/')
    if not u:
        return panel_default_api_url()
    if u == 'dsystem-ar-api.onrender.com':
        return 'https://dsystem-ar-api.onrender.com'
    if u.startswith('http://dsystem-ar-api.onrender.com'):
        return 'https://dsystem-ar-api.onrender.com'
    if re.match(r'^\d+\.\d+\.\d+\.\d+', u):
        return 'http://' + u
    if (('192.168.' in u) or ('10.17.' in u) or ('localhost' in u)) and u.startswith('https://'):
        return u.replace('https://','http://')
    return u

def proxy_json_response(resp):
    try:
        data = resp.json()
    except Exception:
        raise HTTPException(resp.status_code, resp.text[:500])
    if not resp.ok:
        raise HTTPException(resp.status_code, data.get('detail') if isinstance(data, dict) else data)
    return data

@app.post('/api/upload-base-cache')
def upload_base_cache(file: UploadFile=File(...), api_url: str=Form('')):
    if not file.filename.lower().endswith('.xlsx'):
        raise HTTPException(400, 'Envie um arquivo XLSX')

    api = normalize_panel_api_url(api_url)
    try:
        files = {
            'file': (
                file.filename,
                file.file,
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        }
        resp = requests.post(api + '/api/upload-base-cache', files=files, timeout=180)
        return proxy_json_response(resp)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f'Falha ao enviar XLSX para API: {e}')


@app.get('/api/base/sheets')
def base_sheets(api_url: str=''):
    api = normalize_panel_api_url(api_url)
    try:
        resp = requests.get(api + '/api/base/sheets', timeout=60)
        return proxy_json_response(resp)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f'Falha ao consultar abas na API: {e}')


@app.get('/api/base/columns')
def base_columns(sheet: str, api_url: str=''):
    api = normalize_panel_api_url(api_url)
    try:
        resp = requests.get(api + '/api/base/columns', params={'sheet': sheet}, timeout=90)
        return proxy_json_response(resp)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f'Falha ao consultar colunas na API: {e}')


@app.post('/api/import-base')
def import_base(
    sheet: str=Form('TODAS'),
    col_instalacao: str=Form(''),
    col_medidor: str=Form(''),
    col_nome_cliente: str=Form(''),
    api_url: str=Form('')
):
    api = normalize_panel_api_url(api_url)
    try:
        resp = requests.post(api + '/api/import-base', data={
            'sheet': sheet,
            'col_instalacao': col_instalacao,
            'col_medidor': col_medidor,
            'col_nome_cliente': col_nome_cliente
        }, timeout=600)
        return proxy_json_response(resp)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f'Falha ao processar base na API: {e}')



@app.post('/api/base/reindex')
def base_reindex(api_url: str=''):
    api = normalize_panel_api_url(api_url)
    try:
        resp = requests.post(api + '/api/base/reindex', timeout=180)
        return proxy_json_response(resp)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f'Falha ao reindexar base na API: {e}')

@app.get('/api/base/find')
def find_base(instalacao: str='', medidor: str='', nome_cliente: str='', api_url: str=''):
    api = normalize_panel_api_url(api_url)
    try:
        resp = requests.get(api + '/api/base/find', params={
            'instalacao': instalacao,
            'medidor': medidor,
            'nome_cliente': nome_cliente
        }, timeout=60)
        return proxy_json_response(resp)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f'Falha ao consultar base na API: {e}')


@app.post('/api/base/complete/{id}')
def complete_from_base(
    id:int,
    instalacao: str=Form(''),
    medidor: str=Form(''),
    nome_cliente: str=Form(''),
    api_url: str=Form('')
):
    api = normalize_panel_api_url(api_url)

    with conn() as c:
        ar = c.execute('SELECT * FROM ars WHERE id=?',(id,)).fetchone()
        if not ar:
            raise HTTPException(404, 'AR não encontrado')

        search_inst = instalacao or ar['instalacao'] or ''
        search_md = medidor or ar['medidor'] or ''
        search_nome = nome_cliente or ar['nome_cliente'] or ''

    try:
        resp = requests.get(api + '/api/base/find', params={
            'instalacao': search_inst,
            'medidor': search_md,
            'nome_cliente': search_nome
        }, timeout=60)
        r = proxy_json_response(resp)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f'Falha ao consultar base na API: {e}')

    if not r:
        return {'ok': False, 'message': 'Não encontrado na base'}

    with conn() as c:
        ar = c.execute('SELECT * FROM ars WHERE id=?',(id,)).fetchone()
        inst = instalacao or ar['instalacao'] or r.get('instalacao') or ''
        md = medidor or ar['medidor'] or r.get('medidor') or ''
        nome = nome_cliente or ar['nome_cliente'] or r.get('nome_cliente') or ''

        old = UPLOADS / ar['filename']
        ext = old.suffix.lower().replace('.','') or 'pdf'
        newname = build_name(inst, md, ext)
        new = UPLOADS / newname
        base = new.stem
        i = 2
        while new.exists() and new.name != old.name:
            new = UPLOADS / f'{base}_{i}.{ext}'
            i += 1

        if old.exists() and new.name != old.name:
            old.rename(new)

        c.execute(
            'UPDATE ars SET instalacao=?, medidor=?, nome_cliente=?, filename=? WHERE id=?',
            (inst, md, nome, new.name, id)
        )
        c.commit()

    return {'ok': True, 'base': r, 'filename': new.name}

