import os
import re
import json
import tempfile
import threading
import time
import requests as http_requests
import openpyxl
from datetime import date
from flask import Flask, render_template_string, request, jsonify, redirect
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor

# ─── CONFIGURACIÓN ───────────────────────────────────────────────────────────

BASE_URL   = "https://portal.distrinando.com.ar"
COMPANY_ID = "1"

ENDPOINTS = {
    "deporte": {
        "titulo": "Deporte",
        "color": "#6366f1",
        "icon": "&#9917;",
        "from_controller": "MyDocuments",
        "save_from": "MyDocuments",
        "save_url_from": "/Sales/SalesDraft/Index",
        "cuentas": {
            "deporte_cba": {
                "username": "DEPORTE-CBA MAGRIS",
                "password": "1973",
                "label": "Deporte CBA",
                "icon": "&#127939;",
            },
        },
    },
    "moda": {
        "titulo": "Moda",
        "color": "#ec4899",
        "icon": "&#128090;",
        "from_controller": "SalesOrder",
        "save_from": "SalesOrder",
        "save_url_from": "/Sales/SalesOrder/Index",
        "cuentas": {
            "moda_cordoba": {
                "username": "moda_cordoba",
                "password": "1520",
                "label": "Cordoba",
                "icon": "&#127963;",
            },
            "moda_cuyo": {
                "username": "moda_cuyo",
                "password": "1520",
                "label": "Cuyo",
                "icon": "&#9968;",
            },
            "moda_norte": {
                "username": "moda_norte",
                "password": "1520",
                "label": "Norte",
                "icon": "&#9728;",
            },
        },
    },
}

# ─── APP ─────────────────────────────────────────────────────────────────────

app = Flask(__name__)
jobs = {}


# ─── LÓGICA PORTAL (compartida) ─────────────────────────────────────────────

def portal_login(session, username, password):
    r = session.post(
        f"{BASE_URL}/Login/Signin",
        data=f"username={http_requests.utils.quote(username)}&password={password}&rememberMe=undefined",
        headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
    if r.status_code != 200:
        return False
    session.post(
        f"{BASE_URL}/Login/SigninCompany",
        data=f"CompanyId={COMPANY_ID}",
        headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
    return True


def extraer_page_key(session, from_controller):
    form_url = f"{BASE_URL}/Sales/SalesOrder/ActionPurchaseOrder?ActionPurchaseOrder=Add&IdPO=0&fromController={from_controller}"
    form_r = session.get(form_url, headers={"Referer": BASE_URL})
    session.headers.update({"Referer": form_url})

    soup = BeautifulSoup(form_r.text, 'html.parser')
    pk_input = soup.find(id='Pagekey') or soup.find(id='PageKey') or soup.find(id='pageKey')
    if pk_input:
        return pk_input.get('value', '')
    pk_match = re.search(r'id=["\'](?:P|p)age[Kk]ey["\'][^>]*value=["\']([^"\']+)', form_r.text)
    return pk_match.group(1) if pk_match else ''


def init_formulario(session, page_key):
    today = date.today()
    today_fmt = f"{today.month}/{today.day}/{today.year}"
    session.post(f"{BASE_URL}/Sales/SalesOrder/UpdateRateList",
                 data=f"DocDate={http_requests.utils.quote(today_fmt)}&pPageKey={page_key}",
                 headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
    session.post(f"{BASE_URL}/Sales/SalesOrder/_GetDistributionRuleList",
                 headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
    session.post(f"{BASE_URL}/Sales/SalesOrder/_GetGLAccountList",
                 data=f"pPageKey={page_key}",
                 headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
    session.post(f"{BASE_URL}/Sales/SalesOrder/GetItemsModel",
                 headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})


def buscar_cliente(session, nombre):
    payload = (
        "draw=1"
        "&columns[0][data]=CardCode&columns[0][name]=CardCode&columns[0][searchable]=true&columns[0][orderable]=true&columns[0][search][value]=&columns[0][search][regex]=false"
        "&columns[1][data]=CardName&columns[1][name]=CardName&columns[1][searchable]=true&columns[1][orderable]=true&columns[1][search][value]=&columns[1][search][regex]=false"
        "&columns[2][data]=CardType&columns[2][name]=CardType&columns[2][searchable]=true&columns[2][orderable]=true&columns[2][search][value]=&columns[2][search][regex]=false"
        "&columns[3][data]=&columns[3][name]=&columns[3][searchable]=true&columns[3][orderable]=false&columns[3][search][value]=&columns[3][search][regex]=false"
        "&order[0][column]=0&order[0][dir]=asc&start=0&length=10&search[value]=&search[regex]=false"
        f"&pCardCode=&pCardName={http_requests.utils.quote(nombre)}"
    )
    r = session.post(f"{BASE_URL}/Sales/SalesOrder/_Customers", data=payload,
                     headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                              "X-Requested-With": "XMLHttpRequest"})
    data = r.json()
    resultados = data.get("data", [])
    if not resultados:
        return None, None
    c = resultados[0]
    return c["CardCode"], c["CardName"]


def obtener_datos_bp(session, card_code):
    r = session.post(f"{BASE_URL}/Sales/SalesOrder/GetBp",
                     data=f"Id={card_code}&LocalCurrency=ARS",
                     headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                              "X-Requested-With": "XMLHttpRequest"})
    return r.json()


def buscar_item(session, sku, card_code, page_key):
    payload = (
        "draw=1"
        "&columns[0][data]=ItemCode&columns[0][name]=ItemCode&columns[0][searchable]=true&columns[0][orderable]=false&columns[0][search][value]=&columns[0][search][regex]=false"
        "&columns[1][data]=ItemCode&columns[1][name]=ItemCode&columns[1][searchable]=true&columns[1][orderable]=true&columns[1][search][value]=&columns[1][search][regex]=false"
        "&columns[2][data]=ItemName&columns[2][name]=ItemName&columns[2][searchable]=true&columns[2][orderable]=true&columns[2][search][value]=&columns[2][search][regex]=false"
        "&columns[3][data]=&columns[3][name]=Stock&columns[3][searchable]=true&columns[3][orderable]=false&columns[3][search][value]=&columns[3][search][regex]=false"
        "&order[0][column]=0&order[0][dir]=asc&start=0&length=10&search[value]=&search[regex]=false"
        f"&pPageKey={page_key}&pItemCode={http_requests.utils.quote(sku)}"
        f"&pItemName=&pCkCatalogueNum=false&pCardCode={card_code}"
        "&pBPCatalogCode=&pInventoryItem=&pItemWithStock=N&pItemGroup=0"
    )
    r = session.post(f"{BASE_URL}/Sales/SalesOrder/_Items", data=payload,
                     headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                              "X-Requested-With": "XMLHttpRequest"})
    if r.status_code != 200:
        return None
    resultados = r.json().get("data", [])
    if not resultados:
        return None
    for item in resultados:
        if item.get("ItemCode") == sku:
            return item
    return resultados[0]


def consultar_proyecto(session, card_code):
    try:
        qr = session.post(f"{BASE_URL}/QueryManager/GetQueryResult",
                          json={"pQueryIdentifier": "WESAP_TCODE_JUR",
                                "pQueryParams": [{"Key": "Address", "Value": "ENTREGAR EN"},
                                                 {"Key": "CardCode", "Value": card_code}]},
                          headers={"Content-Type": "application/json; charset=UTF-8",
                                   "X-Requested-With": "XMLHttpRequest"})
        if qr.status_code == 200:
            d = json.loads(qr.json())
            if d and d[0].get("PROJECTO"):
                return d[0]["PROJECTO"]
    except Exception:
        pass
    return ''


def consultar_precio(session, item_code, card_code):
    try:
        qr = session.post(f"{BASE_URL}/QueryManager/GetQueryResult",
                          json={"pQueryIdentifier": "qrprecio",
                                "pQueryParams": [{"Key": "ItemCode", "Value": item_code},
                                                 {"Key": "CardCode", "Value": card_code}]},
                          headers={"Content-Type": "application/json; charset=UTF-8",
                                   "X-Requested-With": "XMLHttpRequest"})
        if qr.status_code == 200:
            d = json.loads(qr.json())
            if d and d[0].get("PRECIO"):
                return float(str(d[0]["PRECIO"]).replace(',', '.'))
    except Exception:
        pass
    return None


def agregar_item(session, sku, card_code, page_key, project, log_fn):
    item_search = buscar_item(session, sku, card_code, page_key)
    if not item_search:
        log_fn("    No se encontro en la busqueda")
        return None
    item_code = item_search["ItemCode"]
    item_name = item_search.get("ItemName", item_code)
    log_fn(f"    -> {item_name}")

    with ThreadPoolExecutor(max_workers=2) as pool:
        future_form = pool.submit(
            session.post, f"{BASE_URL}/Sales/SalesOrder/_ItemsForm",
            data=f"pItems={http_requests.utils.quote(item_code)}&pCurrency=ARS&CardCode={card_code}&pPageKey={page_key}&pCkCatalogNum=false",
            headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                     "X-Requested-With": "XMLHttpRequest"})
        future_precio = pool.submit(consultar_precio, session, item_code, card_code)
        r = future_form.result()
        precio_qm = future_precio.result()

    if r.status_code != 200 or 'Value cannot be null' in r.text:
        log_fn(f"    Error en _ItemsForm: status {r.status_code}")
        return None

    try:
        soup = BeautifulSoup(r.text, 'html.parser')

        def get_val(prefix, idx=0):
            el = soup.find(id=f'{prefix}{idx}')
            if el:
                return el.get('value', '') if el.name == 'input' else el.get_text(strip=True)
            return ''

        def to_float(val):
            try: return float(str(val).replace(',', '.'))
            except: return 0.0

        price = precio_qm if precio_qm is not None else to_float(get_val('txt'))
        log_fn(f"    Precio: {price}")

        return {
            'ItemCode': item_code, 'ItemName': get_val('DescItem') or item_name,
            'Price': price, 'PriceBefDi': price,
            'WhsCode': get_val('AutoComplete') or '001', 'OcrCode': '', 'OcrCode2': None,
            'UomCode': get_val('UOMAuto') or '', 'TaxCode': get_val('TaxCode') or 'IVA_21',
            'VatPrcnt': to_float(get_val('VatPrcnt')) or 21.0,
            'DiscPrcnt': to_float(get_val('DiscPrcnt')), 'Project': project,
        }
    except Exception as ex:
        log_fn(f"    Error parseando: {ex}")
        return None


def actualizar_cantidad(session, linea, cantidad, page_key):
    payload = [{
        "Dscription": linea.get("Dscription", linea.get("ItemName", "")),
        "Quantity": cantidad, "Price": linea.get("Price", 0),
        "PriceBefDi": linea.get("PriceBefDi", linea.get("Price", 0)),
        "Currency": "ARS", "LineNum": str(linea.get("LineNum", "0")),
        "WhsCode": linea.get("WhsCode", "001"), "OcrCode": linea.get("OcrCode", ""),
        "OcrCode2": linea.get("OcrCode2", ""), "UomCode": linea.get("UomCode", ""),
        "FreeTxt": "", "GLAccount": {"FormatCode": ""}, "ShipDate": None,
        "TaxCode": linea.get("TaxCode", "IVA_21"), "DiscPrcnt": 0,
        "VatPrcnt": linea.get("VatPrcnt", 21), "MappedUdf": [],
        "SerialBatch": "", "Freight": [{"ExpnsCode": "", "LineTotal": ""}]
    }]
    session.post(f"{BASE_URL}/Sales/SalesOrder/_UpdateLinesChanged?pPageKey={page_key}",
                 json=payload,
                 headers={"Content-Type": "application/json; charset=UTF-8",
                          "X-Requested-With": "XMLHttpRequest"})


def guardar_borrador(session, bp_data, lines, page_key, save_from, save_url_from, comentario=''):
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
        "OwnerCode": "0", "Comments": comentario, "PageKey": page_key,
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
        "MappedUdf": [], "From": save_from, "UrlFrom": save_url_from,
        "QuickOrderId": "", "SaveAsDraft": "true"
    }
    return session.post(f"{BASE_URL}/Sales/SalesOrder/Add", json=body,
                        headers={"Content-Type": "application/json; charset=UTF-8",
                                 "X-Requested-With": "XMLHttpRequest"})


# ─── PROCESO DE CARGA ───────────────────────────────────────────────────────

def correr_carga(job_id, ruta_excel, nombre_cliente, endpoint_key, cuenta_key, descripcion=''):
    job = jobs[job_id]
    log_fn = lambda msg: job["logs"].append(msg)

    try:
        ep = ENDPOINTS[endpoint_key]
        cuenta = ep["cuentas"][cuenta_key]

        session = http_requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
            "X-Requested-With": "XMLHttpRequest",
            "Accept-Language": "es-ES,es;q=0.9",
            "Origin": BASE_URL,
        })

        # Login
        log_fn(f"Iniciando sesion ({cuenta['label']})...")
        if not portal_login(session, cuenta["username"], cuenta["password"]):
            log_fn("Error en login")
            return
        log_fn("Sesion iniciada")

        # Formulario + PageKey
        log_fn("Cargando formulario...")
        page_key = extraer_page_key(session, ep["from_controller"])
        if not page_key:
            log_fn("Error: no se pudo obtener PageKey")
            return
        log_fn(f"PageKey: {page_key}")

        # Cliente
        log_fn(f"Buscando cliente: {nombre_cliente}")
        card_code, card_name = buscar_cliente(session, nombre_cliente)
        if not card_code:
            log_fn(f"No se encontro cliente '{nombre_cliente}'")
            return
        log_fn(f"Cliente: {card_name} ({card_code})")

        bp_data = obtener_datos_bp(session, card_code)
        init_formulario(session, page_key)

        # Excel
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
                log_fn("    Omitido")

        if not lines:
            log_fn("No se pudo agregar ningun item.")
            return

        # Guardar
        log_fn(f"\nGuardando borrador ({len(lines)} items)...")
        r = guardar_borrador(session, bp_data, lines, page_key,
                             ep["save_from"], ep["save_url_from"], descripcion)
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
        try:
            os.unlink(ruta_excel)
        except Exception:
            pass


# ─── TEMPLATES ───────────────────────────────────────────────────────────────

FORM_HTML = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Carga — {{ ep.titulo }}</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            background: #0f172a; color: #e2e8f0;
            min-height: 100vh; display: flex; align-items: center; justify-content: center;
        }
        .container {
            background: #1e293b; border-radius: 16px; padding: 40px;
            width: 520px; box-shadow: 0 25px 50px rgba(0,0,0,0.4);
        }
        h1 { text-align: center; font-size: 1.5rem; font-weight: 600; margin-bottom: 32px; color: #f8fafc; }
        h1 span { color: {{ ep.color }}; }
        .form-group { margin-bottom: 20px; }
        label {
            display: block; font-size: 0.85rem; font-weight: 500; color: #94a3b8;
            margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.5px;
        }
        input[type="text"] {
            width: 100%; padding: 12px 14px; background: #0f172a;
            border: 1px solid #334155; border-radius: 8px; color: #e2e8f0;
            font-size: 0.95rem; outline: none; transition: border-color 0.2s;
        }
        input[type="text"]:focus { border-color: {{ ep.color }}; }
        .account-cards { display: flex; gap: 8px; flex-wrap: wrap; }
        .account-card {
            flex: 1; min-width: 100px; padding: 12px 8px; background: #0f172a;
            border: 2px solid #334155; border-radius: 8px; text-align: center;
            cursor: pointer; transition: all 0.2s; font-size: 0.82rem;
            font-weight: 500; color: #94a3b8;
        }
        .account-card:hover { border-color: {{ ep.color }}; color: #e2e8f0; }
        .account-card.selected { border-color: {{ ep.color }}; background: #1e1b4b; color: #c7d2fe; }
        .account-card .card-icon { font-size: 1.3rem; margin-bottom: 4px; }
        .file-upload {
            position: relative; width: 100%; padding: 24px; background: #0f172a;
            border: 2px dashed #334155; border-radius: 8px; text-align: center;
            cursor: pointer; transition: border-color 0.2s, background 0.2s;
        }
        .file-upload:hover, .file-upload.dragover { border-color: {{ ep.color }}; background: #1a2340; }
        .file-upload.has-file { border-color: #34d399; border-style: solid; }
        .file-upload input[type="file"] { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
        .file-upload .icon { font-size: 1.6rem; margin-bottom: 6px; }
        .file-upload .text { font-size: 0.85rem; color: #94a3b8; }
        .file-upload .filename { font-size: 0.9rem; color: #34d399; font-weight: 500; display: none; }
        .btn {
            width: 100%; padding: 14px; background: {{ ep.color }}; color: white;
            border: none; border-radius: 8px; font-size: 1rem; font-weight: 600;
            cursor: pointer; transition: all 0.2s; margin-top: 8px;
        }
        .btn:hover:not(:disabled) { filter: brightness(1.2); transform: translateY(-1px); }
        .btn:disabled { background: #334155; color: #64748b; cursor: not-allowed; }
        .log-container { margin-top: 24px; display: none; }
        .log-container.visible { display: block; }
        .log-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }
        .log-header label { margin-bottom: 0; }
        .spinner {
            width: 18px; height: 18px; border: 2px solid #334155;
            border-top-color: {{ ep.color }}; border-radius: 50%;
            animation: spin 0.8s linear infinite; display: none;
        }
        .spinner.active { display: inline-block; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .log {
            background: #0f172a; border: 1px solid #334155; border-radius: 8px;
            padding: 14px; height: 260px; overflow-y: auto;
            font-family: 'Cascadia Code', 'Fira Code', 'Consolas', monospace;
            font-size: 0.8rem; line-height: 1.6; white-space: pre-wrap; word-break: break-word;
        }
        .log::-webkit-scrollbar { width: 6px; }
        .log::-webkit-scrollbar-track { background: transparent; }
        .log::-webkit-scrollbar-thumb { background: #334155; border-radius: 3px; }
        .status-bar {
            margin-top: 12px; padding: 10px 14px; border-radius: 8px;
            font-size: 0.85rem; font-weight: 500; display: none; text-align: center;
        }
        .status-bar.success { display: block; background: #064e3b; color: #34d399; border: 1px solid #065f46; }
        .status-bar.error { display: block; background: #450a0a; color: #fca5a5; border: 1px solid #7f1d1d; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Carga de Pedidos <span>{{ ep.titulo }}</span></h1>

        <div class="form-group">
            <label>Cuenta</label>
            <div class="account-cards">
                {% for key, c in ep.cuentas.items() %}
                <div class="account-card" data-value="{{ key }}" onclick="selectAccount(this)">
                    <div class="card-icon">{{ c.icon | safe }}</div>
                    {{ c.label }}
                </div>
                {% endfor %}
            </div>
            <input type="hidden" id="cuenta" value="">
        </div>

        <div class="form-group">
            <label>Nombre del cliente</label>
            <input type="text" id="cliente" placeholder="Ej: AVENDANO">
        </div>

        <div class="form-group">
            <label>Descripcion / Comentario</label>
            <input type="text" id="descripcion" placeholder="Opcional">
        </div>

        <div class="form-group">
            <label>Archivo Excel (SKU / Cantidad)</label>
            <div class="file-upload" id="dropZone">
                <div class="icon">&#128196;</div>
                <div class="text">Arrastra o hace click para seleccionar</div>
                <div class="filename" id="fileName"></div>
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
        const ENDPOINT = '{{ endpoint_key }}';
        let selectedAccount = '';

        function selectAccount(el) {
            document.querySelectorAll('.account-card').forEach(c => c.classList.remove('selected'));
            el.classList.add('selected');
            selectedAccount = el.dataset.value;
            document.getElementById('cuenta').value = selectedAccount;
        }

        // Auto-select if only one account
        const cards = document.querySelectorAll('.account-card');
        if (cards.length === 1) selectAccount(cards[0]);

        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const fileNameEl = document.getElementById('fileName');

        ['dragover','dragenter'].forEach(ev =>
            dropZone.addEventListener(ev, e => { e.preventDefault(); dropZone.classList.add('dragover'); }));
        ['dragleave','drop'].forEach(ev =>
            dropZone.addEventListener(ev, () => dropZone.classList.remove('dragover')));
        dropZone.addEventListener('drop', e => {
            e.preventDefault();
            if (e.dataTransfer.files.length) { fileInput.files = e.dataTransfer.files; showFile(e.dataTransfer.files[0].name); }
        });
        fileInput.addEventListener('change', () => { if (fileInput.files.length) showFile(fileInput.files[0].name); });

        function showFile(name) {
            dropZone.querySelector('.icon').style.display = 'none';
            dropZone.querySelector('.text').style.display = 'none';
            fileNameEl.textContent = name; fileNameEl.style.display = 'block';
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
            btn.disabled = true; btn.textContent = 'Cargando...';
            const logBox = document.getElementById('logBox');
            logBox.textContent = '';
            document.getElementById('logContainer').classList.add('visible');
            document.getElementById('spinner').classList.add('active');
            const statusBar = document.getElementById('statusBar');
            statusBar.className = 'status-bar'; statusBar.style.display = 'none';

            const descripcion = document.getElementById('descripcion').value.trim();
            const fd = new FormData();
            fd.append('endpoint', ENDPOINT);
            fd.append('cuenta', cuenta);
            fd.append('cliente', cliente);
            fd.append('descripcion', descripcion);
            fd.append('file', file);

            fetch('/cargar', { method: 'POST', body: fd })
                .then(r => r.json()).then(data => { if (data.job_id) pollLogs(data.job_id); })
                .catch(err => { logBox.textContent += 'Error: ' + err + '\\n'; btn.disabled = false; btn.textContent = 'Iniciar carga'; });
        }

        function pollLogs(jobId) {
            const logBox = document.getElementById('logBox');
            const btn = document.getElementById('btnCargar');
            let idx = 0;
            const interval = setInterval(() => {
                fetch('/logs/' + jobId + '?from=' + idx).then(r => r.json()).then(data => {
                    data.logs.forEach(msg => { logBox.textContent += msg + '\\n'; });
                    idx += data.logs.length;
                    logBox.scrollTop = logBox.scrollHeight;
                    if (data.done) {
                        clearInterval(interval);
                        document.getElementById('spinner').classList.remove('active');
                        btn.disabled = false; btn.textContent = 'Iniciar carga';
                        const sb = document.getElementById('statusBar');
                        if (data.success) { sb.textContent = 'Borrador guardado exitosamente'; sb.className = 'status-bar success'; }
                        else { sb.textContent = 'Error al procesar el pedido'; sb.className = 'status-bar error'; }
                    }
                });
            }, 500);
        }
    </script>
</body>
</html>
"""


# ─── RUTAS ───────────────────────────────────────────────────────────────────

@app.route('/deportgm')
def deportgm():
    return render_template_string(FORM_HTML, ep=ENDPOINTS["deporte"], endpoint_key="deporte")


@app.route('/moda')
def moda():
    return render_template_string(FORM_HTML, ep=ENDPOINTS["moda"], endpoint_key="moda")


@app.route('/cargar', methods=['POST'])
def cargar():
    endpoint_key = request.form.get('endpoint')
    cuenta_key = request.form.get('cuenta')
    cliente = request.form.get('cliente', '').strip()
    descripcion = request.form.get('descripcion', '').strip()
    file = request.files.get('file')

    if endpoint_key not in ENDPOINTS:
        return jsonify({"error": "Endpoint invalido"}), 400
    if cuenta_key not in ENDPOINTS[endpoint_key]["cuentas"]:
        return jsonify({"error": "Cuenta invalida"}), 400
    if not cliente:
        return jsonify({"error": "Falta cliente"}), 400
    if not file:
        return jsonify({"error": "Falta archivo"}), 400

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    file.save(tmp.name)
    tmp.close()

    job_id = str(int(time.time() * 1000))
    jobs[job_id] = {"logs": [], "done": False, "success": False}

    threading.Thread(
        target=correr_carga,
        args=(job_id, tmp.name, cliente, endpoint_key, cuenta_key, descripcion),
        daemon=True
    ).start()

    return jsonify({"job_id": job_id})


@app.route('/logs/<job_id>')
def get_logs(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"logs": [], "done": True, "success": False})
    from_idx = int(request.args.get('from', 0))
    return jsonify({"logs": job["logs"][from_idx:], "done": job["done"], "success": job["success"]})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"Servidor corriendo en http://localhost:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)
