import os
import json
import tempfile
import threading
import queue
import time
import requests
import openpyxl
from datetime import date
from flask import Flask, render_template_string, request, jsonify
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor

# ─── CONFIGURACIÓN ───────────────────────────────────────────────────────────

BASE_URL   = "https://portal.distrinando.com.ar"
COMPANY_ID = "1"

CUENTAS = {
    "moda_cordoba": {"username": "moda_cordoba", "label": "Moda Cordoba"},
    "moda_cuyo":    {"username": "moda_cuyo",    "label": "Moda Cuyo"},
    "moda_norte":   {"username": "moda_norte",   "label": "Moda Norte"},
}
PASSWORD = "1520"

# ─── APP FLASK ───────────────────────────────────────────────────────────────

app = Flask(__name__)

# Cola de logs por job
jobs = {}

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Carga de Pedidos — Moda</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            background: #0f172a;
            color: #e2e8f0;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .container {
            background: #1e293b;
            border-radius: 16px;
            padding: 40px;
            width: 520px;
            box-shadow: 0 25px 50px rgba(0,0,0,0.4);
        }
        h1 {
            text-align: center;
            font-size: 1.5rem;
            font-weight: 600;
            margin-bottom: 32px;
            color: #f8fafc;
        }
        h1 span { color: #818cf8; }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            font-size: 0.85rem;
            font-weight: 500;
            color: #94a3b8;
            margin-bottom: 6px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        select, input[type="text"] {
            width: 100%;
            padding: 12px 14px;
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 8px;
            color: #e2e8f0;
            font-size: 0.95rem;
            transition: border-color 0.2s;
            outline: none;
        }
        select:focus, input[type="text"]:focus {
            border-color: #818cf8;
        }
        select option { background: #1e293b; }
        .file-upload {
            position: relative;
            width: 100%;
            padding: 24px;
            background: #0f172a;
            border: 2px dashed #334155;
            border-radius: 8px;
            text-align: center;
            cursor: pointer;
            transition: border-color 0.2s, background 0.2s;
        }
        .file-upload:hover, .file-upload.dragover {
            border-color: #818cf8;
            background: #1a2340;
        }
        .file-upload.has-file {
            border-color: #34d399;
            border-style: solid;
        }
        .file-upload input[type="file"] {
            position: absolute;
            inset: 0;
            opacity: 0;
            cursor: pointer;
        }
        .file-upload .icon { font-size: 1.6rem; margin-bottom: 6px; }
        .file-upload .text { font-size: 0.85rem; color: #94a3b8; }
        .file-upload .filename {
            font-size: 0.9rem;
            color: #34d399;
            font-weight: 500;
        }
        .btn {
            width: 100%;
            padding: 14px;
            background: #6366f1;
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: background 0.2s, transform 0.1s;
            margin-top: 8px;
        }
        .btn:hover:not(:disabled) { background: #818cf8; transform: translateY(-1px); }
        .btn:active:not(:disabled) { transform: translateY(0); }
        .btn:disabled {
            background: #334155;
            color: #64748b;
            cursor: not-allowed;
        }
        .log-container {
            margin-top: 24px;
            display: none;
        }
        .log-container.visible { display: block; }
        .log-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 8px;
        }
        .log-header label { margin-bottom: 0; }
        .spinner {
            width: 18px; height: 18px;
            border: 2px solid #334155;
            border-top-color: #818cf8;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            display: none;
        }
        .spinner.active { display: inline-block; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .log {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 8px;
            padding: 14px;
            height: 260px;
            overflow-y: auto;
            font-family: 'Cascadia Code', 'Fira Code', 'Consolas', monospace;
            font-size: 0.8rem;
            line-height: 1.6;
            white-space: pre-wrap;
            word-break: break-word;
        }
        .log::-webkit-scrollbar { width: 6px; }
        .log::-webkit-scrollbar-track { background: transparent; }
        .log::-webkit-scrollbar-thumb { background: #334155; border-radius: 3px; }
        .status-bar {
            margin-top: 12px;
            padding: 10px 14px;
            border-radius: 8px;
            font-size: 0.85rem;
            font-weight: 500;
            display: none;
            text-align: center;
        }
        .status-bar.success { display: block; background: #064e3b; color: #34d399; border: 1px solid #065f46; }
        .status-bar.error { display: block; background: #450a0a; color: #fca5a5; border: 1px solid #7f1d1d; }
        .account-cards {
            display: flex;
            gap: 8px;
        }
        .account-card {
            flex: 1;
            padding: 12px 8px;
            background: #0f172a;
            border: 2px solid #334155;
            border-radius: 8px;
            text-align: center;
            cursor: pointer;
            transition: all 0.2s;
            font-size: 0.82rem;
            font-weight: 500;
            color: #94a3b8;
        }
        .account-card:hover { border-color: #818cf8; color: #e2e8f0; }
        .account-card.selected {
            border-color: #818cf8;
            background: #1e1b4b;
            color: #c7d2fe;
        }
        .account-card .card-icon { font-size: 1.3rem; margin-bottom: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Carga de Pedidos <span>Moda</span></h1>

        <div class="form-group">
            <label>Cuenta</label>
            <div class="account-cards">
                <div class="account-card" data-value="moda_cordoba" onclick="selectAccount(this)">
                    <div class="card-icon">&#127963;</div>
                    Cordoba
                </div>
                <div class="account-card" data-value="moda_cuyo" onclick="selectAccount(this)">
                    <div class="card-icon">&#9968;</div>
                    Cuyo
                </div>
                <div class="account-card" data-value="moda_norte" onclick="selectAccount(this)">
                    <div class="card-icon">&#9728;</div>
                    Norte
                </div>
            </div>
            <input type="hidden" id="cuenta" value="">
        </div>

        <div class="form-group">
            <label>Nombre del cliente</label>
            <input type="text" id="cliente" placeholder="Ej: AVENDANO">
        </div>

        <div class="form-group">
            <label>Archivo Excel (SKU / Cantidad)</label>
            <div class="file-upload" id="dropZone">
                <div class="icon">&#128196;</div>
                <div class="text">Arrastra un archivo o hace click para seleccionar</div>
                <div class="filename" id="fileName" style="display:none"></div>
                <input type="file" id="fileInput" accept=".xlsx,.xls">
            </div>
        </div>

        <button class="btn" id="btnCargar" onclick="iniciarCarga()">Iniciar carga</button>

        <div class="log-container" id="logContainer">
            <div class="log-header">
                <label>Log</label>
                <div class="spinner" id="spinner"></div>
            </div>
            <div class="log" id="logBox"></div>
        </div>
        <div class="status-bar" id="statusBar"></div>
    </div>

    <script>
        let selectedAccount = '';

        function selectAccount(el) {
            document.querySelectorAll('.account-card').forEach(c => c.classList.remove('selected'));
            el.classList.add('selected');
            selectedAccount = el.dataset.value;
            document.getElementById('cuenta').value = selectedAccount;
        }

        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const fileName = document.getElementById('fileName');

        ['dragover','dragenter'].forEach(ev => {
            dropZone.addEventListener(ev, e => { e.preventDefault(); dropZone.classList.add('dragover'); });
        });
        ['dragleave','drop'].forEach(ev => {
            dropZone.addEventListener(ev, () => dropZone.classList.remove('dragover'));
        });
        dropZone.addEventListener('drop', e => {
            e.preventDefault();
            if (e.dataTransfer.files.length) {
                fileInput.files = e.dataTransfer.files;
                showFile(e.dataTransfer.files[0].name);
            }
        });
        fileInput.addEventListener('change', () => {
            if (fileInput.files.length) showFile(fileInput.files[0].name);
        });

        function showFile(name) {
            dropZone.querySelector('.icon').style.display = 'none';
            dropZone.querySelector('.text').style.display = 'none';
            fileName.textContent = name;
            fileName.style.display = 'block';
            dropZone.classList.add('has-file');
        }

        function iniciarCarga() {
            const cuenta = document.getElementById('cuenta').value;
            const cliente = document.getElementById('cliente').value.trim();
            const file = fileInput.files[0];

            if (!cuenta) return alert('Selecciona una cuenta.');
            if (!cliente) return alert('Ingresa el nombre del cliente.');
            if (!file) return alert('Selecciona un archivo Excel.');

            const btn = document.getElementById('btnCargar');
            btn.disabled = true;
            btn.textContent = 'Cargando...';

            const logBox = document.getElementById('logBox');
            logBox.textContent = '';
            document.getElementById('logContainer').classList.add('visible');
            document.getElementById('spinner').classList.add('active');
            const statusBar = document.getElementById('statusBar');
            statusBar.className = 'status-bar';
            statusBar.style.display = 'none';

            const formData = new FormData();
            formData.append('cuenta', cuenta);
            formData.append('cliente', cliente);
            formData.append('file', file);

            fetch('/cargar', { method: 'POST', body: formData })
                .then(r => r.json())
                .then(data => {
                    if (data.job_id) pollLogs(data.job_id);
                })
                .catch(err => {
                    logBox.textContent += 'Error de conexion: ' + err + '\\n';
                    btn.disabled = false;
                    btn.textContent = 'Iniciar carga';
                });
        }

        function pollLogs(jobId) {
            const logBox = document.getElementById('logBox');
            const btn = document.getElementById('btnCargar');
            let idx = 0;

            const interval = setInterval(() => {
                fetch('/logs/' + jobId + '?from=' + idx)
                    .then(r => r.json())
                    .then(data => {
                        data.logs.forEach(msg => {
                            logBox.textContent += msg + '\\n';
                        });
                        idx += data.logs.length;
                        logBox.scrollTop = logBox.scrollHeight;

                        if (data.done) {
                            clearInterval(interval);
                            document.getElementById('spinner').classList.remove('active');
                            btn.disabled = false;
                            btn.textContent = 'Iniciar carga';

                            const statusBar = document.getElementById('statusBar');
                            if (data.success) {
                                statusBar.textContent = 'Borrador guardado exitosamente';
                                statusBar.className = 'status-bar success';
                            } else {
                                statusBar.textContent = 'Error al procesar el pedido';
                                statusBar.className = 'status-bar error';
                            }
                        }
                    });
            }, 500);
        }
    </script>
</body>
</html>
"""


# ─── LÓGICA DE CARGA ────────────────────────────────────────────────────────

def buscar_cliente(session, nombre):
    payload = (
        "draw=1"
        "&columns[0][data]=CardCode&columns[0][name]=CardCode&columns[0][searchable]=true&columns[0][orderable]=true&columns[0][search][value]=&columns[0][search][regex]=false"
        "&columns[1][data]=CardName&columns[1][name]=CardName&columns[1][searchable]=true&columns[1][orderable]=true&columns[1][search][value]=&columns[1][search][regex]=false"
        "&columns[2][data]=CardType&columns[2][name]=CardType&columns[2][searchable]=true&columns[2][orderable]=true&columns[2][search][value]=&columns[2][search][regex]=false"
        "&columns[3][data]=&columns[3][name]=&columns[3][searchable]=true&columns[3][orderable]=false&columns[3][search][value]=&columns[3][search][regex]=false"
        "&order[0][column]=0&order[0][dir]=asc&start=0&length=10&search[value]=&search[regex]=false"
        f"&pCardCode=&pCardName={requests.utils.quote(nombre)}"
    )
    r = session.post(
        f"{BASE_URL}/Sales/SalesOrder/_Customers",
        data=payload,
        headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                 "X-Requested-With": "XMLHttpRequest"}
    )
    data = r.json()
    resultados = data.get("data", [])
    if not resultados:
        return None, None
    cliente = resultados[0]
    return cliente["CardCode"], cliente["CardName"]


def obtener_datos_bp(session, card_code):
    r = session.post(
        f"{BASE_URL}/Sales/SalesOrder/GetBp",
        data=f"Id={card_code}&LocalCurrency=ARS",
        headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                 "X-Requested-With": "XMLHttpRequest"}
    )
    return r.json()


def buscar_item(session, sku, card_code, page_key):
    payload = (
        "draw=1"
        "&columns[0][data]=ItemCode&columns[0][name]=ItemCode&columns[0][searchable]=true&columns[0][orderable]=false&columns[0][search][value]=&columns[0][search][regex]=false"
        "&columns[1][data]=ItemCode&columns[1][name]=ItemCode&columns[1][searchable]=true&columns[1][orderable]=true&columns[1][search][value]=&columns[1][search][regex]=false"
        "&columns[2][data]=ItemName&columns[2][name]=ItemName&columns[2][searchable]=true&columns[2][orderable]=true&columns[2][search][value]=&columns[2][search][regex]=false"
        "&columns[3][data]=&columns[3][name]=Stock&columns[3][searchable]=true&columns[3][orderable]=false&columns[3][search][value]=&columns[3][search][regex]=false"
        "&order[0][column]=0&order[0][dir]=asc&start=0&length=10&search[value]=&search[regex]=false"
        f"&pPageKey={page_key}"
        f"&pItemCode={requests.utils.quote(sku)}"
        "&pItemName=&pCkCatalogueNum=false"
        f"&pCardCode={card_code}"
        "&pBPCatalogCode=&pInventoryItem=&pItemWithStock=N&pItemGroup=0"
    )
    r = session.post(
        f"{BASE_URL}/Sales/SalesOrder/_Items",
        data=payload,
        headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                 "X-Requested-With": "XMLHttpRequest"}
    )
    if r.status_code != 200:
        return None
    data = r.json()
    resultados = data.get("data", [])
    if not resultados:
        return None
    for item in resultados:
        if item.get("ItemCode") == sku:
            return item
    return resultados[0]


def consultar_proyecto(session, card_code):
    try:
        qr = session.post(
            f"{BASE_URL}/QueryManager/GetQueryResult",
            json={"pQueryIdentifier": "WESAP_TCODE_JUR",
                  "pQueryParams": [{"Key": "Address", "Value": "ENTREGAR EN"},
                                   {"Key": "CardCode", "Value": card_code}]},
            headers={"Content-Type": "application/json; charset=UTF-8",
                     "X-Requested-With": "XMLHttpRequest"}
        )
        if qr.status_code == 200:
            proj_data = json.loads(qr.json())
            if proj_data and proj_data[0].get("PROJECTO"):
                return proj_data[0]["PROJECTO"]
    except Exception:
        pass
    return ''


def consultar_precio(session, item_code, card_code):
    try:
        qr = session.post(
            f"{BASE_URL}/QueryManager/GetQueryResult",
            json={"pQueryIdentifier": "qrprecio",
                  "pQueryParams": [{"Key": "ItemCode", "Value": item_code},
                                   {"Key": "CardCode", "Value": card_code}]},
            headers={"Content-Type": "application/json; charset=UTF-8",
                     "X-Requested-With": "XMLHttpRequest"}
        )
        if qr.status_code == 200:
            precio_data = json.loads(qr.json())
            if precio_data and precio_data[0].get("PRECIO"):
                return float(str(precio_data[0]["PRECIO"]).replace(',', '.'))
    except Exception:
        pass
    return None


def agregar_item(session, sku, card_code, page_key, project, log_fn):
    item_search = buscar_item(session, sku, card_code, page_key)
    if not item_search:
        log_fn(f"    No se encontro en la busqueda")
        return None
    item_code = item_search["ItemCode"]
    item_name = item_search.get("ItemName", item_code)
    log_fn(f"    -> {item_name}")

    with ThreadPoolExecutor(max_workers=2) as pool:
        future_form = pool.submit(
            session.post,
            f"{BASE_URL}/Sales/SalesOrder/_ItemsForm",
            data=f"pItems={requests.utils.quote(item_code)}&pCurrency=ARS&CardCode={card_code}&pPageKey={page_key}&pCkCatalogNum=false",
            headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                     "X-Requested-With": "XMLHttpRequest"}
        )
        future_precio = pool.submit(consultar_precio, session, item_code, card_code)
        r = future_form.result()
        precio_qm = future_precio.result()

    if r.status_code != 200 or 'Value cannot be null' in r.text:
        log_fn(f"    Error en _ItemsForm: status {r.status_code}")
        return None

    try:
        soup = BeautifulSoup(r.text, 'html.parser')

        def get_input_val(prefix, idx=0):
            el = soup.find(id=f'{prefix}{idx}')
            if el:
                return el.get('value', '') if el.name == 'input' else el.get_text(strip=True)
            return ''

        def to_float(val):
            try:
                return float(str(val).replace(',', '.'))
            except:
                return 0.0

        price = precio_qm if precio_qm is not None else to_float(get_input_val('txt'))
        log_fn(f"    Precio: {price}")

        return {
            'ItemCode': item_code,
            'ItemName': get_input_val('DescItem') or item_name,
            'Price': price,
            'PriceBefDi': price,
            'WhsCode': get_input_val('AutoComplete') or '001',
            'OcrCode': '',
            'OcrCode2': None,
            'UomCode': get_input_val('UOMAuto') or '',
            'TaxCode': get_input_val('TaxCode') or 'IVA_21',
            'VatPrcnt': to_float(get_input_val('VatPrcnt')) or 21.0,
            'DiscPrcnt': to_float(get_input_val('DiscPrcnt')),
            'Project': project,
        }
    except Exception as ex:
        log_fn(f"    Error parseando: {ex}")
        return None


def actualizar_cantidad(session, linea, cantidad, page_key):
    payload = [{
        "Dscription": linea.get("Dscription", linea.get("ItemName", "")),
        "Quantity": cantidad,
        "Price": linea.get("Price", 0),
        "PriceBefDi": linea.get("PriceBefDi", linea.get("Price", 0)),
        "Currency": "ARS",
        "LineNum": str(linea.get("LineNum", "0")),
        "WhsCode": linea.get("WhsCode", "001"),
        "OcrCode": linea.get("OcrCode", ""),
        "OcrCode2": linea.get("OcrCode2", ""),
        "UomCode": linea.get("UomCode", ""),
        "FreeTxt": "",
        "GLAccount": {"FormatCode": ""},
        "ShipDate": None,
        "TaxCode": linea.get("TaxCode", "IVA_21"),
        "DiscPrcnt": 0,
        "VatPrcnt": linea.get("VatPrcnt", 21),
        "MappedUdf": [],
        "SerialBatch": "",
        "Freight": [{"ExpnsCode": "", "LineTotal": ""}]
    }]
    session.post(
        f"{BASE_URL}/Sales/SalesOrder/_UpdateLinesChanged?pPageKey={page_key}",
        json=payload,
        headers={"Content-Type": "application/json; charset=UTF-8",
                 "X-Requested-With": "XMLHttpRequest"}
    )


def guardar_borrador(session, bp_data, lines, page_key):
    today = date.today()
    today_str = f"{today.month}/{today.day}/{today.year}"
    ship_addr = next((a for a in bp_data.get("Addresses", []) if a.get("AdresType") == "S"), {})
    bill_addr = next((a for a in bp_data.get("Addresses", []) if a.get("AdresType") == "B"), {})

    body = {
        "DocDate": today_str, "DocDueDate": today_str, "TaxDate": today_str,
        "CardCode": bp_data["CardCode"], "DocCur": "ARS", "DocRate": "1",
        "DocTotal": sum(l.get("Price", 0) * l.get("Quantity", 1) for l in lines),
        "CardName": bp_data["CardName"], "DiscPrcnt": "0",
        "CntctCode": str(bp_data.get("ListContact", [{}])[0].get("CntctCode", "")),
        "SlpCode": str(bp_data.get("SlpCode", "0")),
        "TrnspCode": None, "NumAtCard": "", "CancelDate": None, "ReqDate": None,
        "OwnerCode": "0", "Comments": "", "PageKey": page_key,
        "SOAddress": {
            "DocEntry": "0",
            "StreetS": ship_addr.get("Street", ""), "StreetB": bill_addr.get("Street", ""),
            "StreetNoS": "", "StreetNoB": "", "BlockS": "", "BlockB": "",
            "CityS": ship_addr.get("City", ""), "CityB": bill_addr.get("City", ""),
            "ZipCodeS": ship_addr.get("ZipCode", ""), "ZipCodeB": bill_addr.get("ZipCode", ""),
            "CountyS": "", "CountyB": "", "StateS": None, "StateB": None,
            "CountryS": None, "CountryB": None, "BuildingS": "", "BuildingB": "",
            "GlbLocNumS": "", "GlbLocNumB": ""
        },
        "ShipToCode": bp_data.get("ShipToDef", "ENTREGAR EN"),
        "PayToCode": bp_data.get("BillToDef", "FACTURAR A"),
        "GroupNum": str(bp_data.get("ListNum", "")),
        "PeyMethod": "", "ListItem": [], "Lines": lines, "ListFreight": [],
        "MappedUdf": [], "From": "SalesOrder", "UrlFrom": "/Sales/SalesOrder/Index",
        "QuickOrderId": "", "SaveAsDraft": "true"
    }
    return session.post(
        f"{BASE_URL}/Sales/SalesOrder/Add",
        json=body,
        headers={"Content-Type": "application/json; charset=UTF-8",
                 "X-Requested-With": "XMLHttpRequest"}
    )


def correr_carga(job_id, ruta_excel, nombre_cliente, cuenta_key):
    job = jobs[job_id]
    log_fn = lambda msg: job["logs"].append(msg)

    try:
        cuenta = CUENTAS[cuenta_key]
        username = cuenta["username"]

        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
            "X-Requested-With": "XMLHttpRequest",
            "Accept-Language": "es-ES,es;q=0.9",
            "Origin": BASE_URL,
        })

        # Login
        log_fn(f"Iniciando sesion ({cuenta['label']})...")
        r = session.post(
            f"{BASE_URL}/Login/Signin",
            data=f"username={requests.utils.quote(username)}&password={PASSWORD}&rememberMe=undefined",
            headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
        )
        if r.status_code != 200:
            log_fn(f"Error en login: {r.status_code}")
            return

        session.post(
            f"{BASE_URL}/Login/SigninCompany",
            data=f"CompanyId={COMPANY_ID}",
            headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
        )
        log_fn("Sesion iniciada")

        # Cargar formulario y extraer PageKey
        FORM_URL = f"{BASE_URL}/Sales/SalesOrder/ActionPurchaseOrder?ActionPurchaseOrder=Add&IdPO=0&fromController=SalesOrder"
        log_fn("Cargando formulario...")
        form_r = session.get(FORM_URL, headers={"Referer": BASE_URL})
        session.headers.update({"Referer": FORM_URL})

        form_soup = BeautifulSoup(form_r.text, 'html.parser')
        pk_input = form_soup.find(id='Pagekey') or form_soup.find(id='PageKey') or form_soup.find(id='pageKey')
        if pk_input:
            page_key = pk_input.get('value', '')
        else:
            import re
            pk_match = re.search(r'id=["\'](?:P|p)age[Kk]ey["\'][^>]*value=["\']([^"\']+)', form_r.text)
            page_key = pk_match.group(1) if pk_match else ''

        if not page_key:
            log_fn("Error: no se pudo obtener PageKey del servidor")
            return

        # Buscar cliente
        log_fn(f"Buscando cliente: {nombre_cliente}")
        card_code, card_name = buscar_cliente(session, nombre_cliente)
        if not card_code:
            log_fn(f"No se encontro cliente '{nombre_cliente}'")
            return
        log_fn(f"Cliente: {card_name} ({card_code})")

        bp_data = obtener_datos_bp(session, card_code)

        # Inicializar formulario
        log_fn(f"PageKey: {page_key}")
        today = date.today()
        today_fmt = f"{today.month}/{today.day}/{today.year}"
        session.post(f"{BASE_URL}/Sales/SalesOrder/UpdateRateList",
                     data=f"DocDate={requests.utils.quote(today_fmt)}&pPageKey={page_key}",
                     headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
        session.post(f"{BASE_URL}/Sales/SalesOrder/_GetDistributionRuleList",
                     headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
        session.post(f"{BASE_URL}/Sales/SalesOrder/_GetGLAccountList",
                     data=f"pPageKey={page_key}",
                     headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
        session.post(f"{BASE_URL}/Sales/SalesOrder/GetItemsModel",
                     headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})

        # Leer Excel
        log_fn("Leyendo Excel...")
        wb = openpyxl.load_workbook(ruta_excel)
        ws = wb.active
        items_excel = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            sku = str(row[0]).strip() if row[0] else None
            cantidad = int(row[1]) if row[1] else 1
            if sku:
                items_excel.append((sku, cantidad))
        log_fn(f"{len(items_excel)} items encontrados")

        # Proyecto
        project = consultar_proyecto(session, card_code)
        if project:
            log_fn(f"Proyecto: {project}")

        # Agregar items
        lines = []
        for idx, (sku, cantidad) in enumerate(items_excel):
            log_fn(f"[{idx+1}/{len(items_excel)}] {sku} x{cantidad}")
            item_data = agregar_item(session, sku, card_code, page_key, project, log_fn)
            if item_data:
                item_data["Quantity"] = cantidad
                item_data["LineNum"] = str(idx)
                actualizar_cantidad(session, item_data, cantidad, page_key)
                lines.append({
                    "ItemCode": item_data.get("ItemCode", sku),
                    "Dscription": item_data.get("ItemName", sku),
                    "Quantity": cantidad,
                    "Price": item_data.get("Price", 0),
                    "PriceBefDi": item_data.get("Price", 0),
                    "Currency": "ARS", "LineNum": str(idx),
                    "WhsCode": item_data.get("WhsCode", "001"),
                    "OcrCode": item_data.get("OcrCode", ""),
                    "OcrCode2": item_data.get("OcrCode2", None),
                    "BaseType": "", "BaseEntry": "", "BaseLine": "",
                    "FreeTxt": "", "GLAccount": {"FormatCode": ""},
                    "Project": item_data.get("Project", ""),
                    "UomCode": item_data.get("UomCode", ""),
                    "SerialBatch": "", "ShipDate": None, "MappedUdf": [],
                    "TaxCode": item_data.get("TaxCode", "IVA_21"),
                    "DiscPrcnt": str(item_data.get("DiscPrcnt", "0.0000")),
                    "VatPrcnt": str(item_data.get("VatPrcnt", "21.0000")),
                    "Freight": [{"ExpnsCode": "", "LineTotal": ""}]
                })
            else:
                log_fn(f"    Omitido")

        if not lines:
            log_fn("No se pudo agregar ningun item.")
            return

        # Guardar
        log_fn(f"\nGuardando borrador ({len(lines)} items)...")
        r = guardar_borrador(session, bp_data, lines, page_key)

        if r.status_code == 200:
            log_fn("Borrador guardado exitosamente!")
            job["success"] = True
        else:
            log_fn(f"Error al guardar. Status: {r.status_code}")
            log_fn(r.text[:300])

    except Exception as ex:
        log_fn(f"Error: {ex}")
        import traceback
        log_fn(traceback.format_exc())
    finally:
        job["done"] = True


# ─── RUTAS ───────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/cargar', methods=['POST'])
def cargar():
    cuenta = request.form.get('cuenta')
    cliente = request.form.get('cliente', '').strip()
    file = request.files.get('file')

    if not cuenta or cuenta not in CUENTAS:
        return jsonify({"error": "Cuenta invalida"}), 400
    if not cliente:
        return jsonify({"error": "Falta cliente"}), 400
    if not file:
        return jsonify({"error": "Falta archivo"}), 400

    # Guardar archivo temporal
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    file.save(tmp.name)
    tmp.close()

    job_id = str(int(time.time() * 1000))
    jobs[job_id] = {"logs": [], "done": False, "success": False}

    threading.Thread(
        target=correr_carga,
        args=(job_id, tmp.name, cliente, cuenta),
        daemon=True
    ).start()

    return jsonify({"job_id": job_id})


@app.route('/logs/<job_id>')
def get_logs(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"logs": [], "done": True, "success": False})

    from_idx = int(request.args.get('from', 0))
    logs = job["logs"][from_idx:]
    return jsonify({"logs": logs, "done": job["done"], "success": job["success"]})


if __name__ == '__main__':
    print("Servidor corriendo en http://localhost:5000")
    app.run(debug=False, port=5000)
