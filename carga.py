import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import requests
import openpyxl
from datetime import date
import random
import string
import json
import threading

# ─── CONFIGURACIÓN (modificar si cambian las credenciales) ───────────────────

BASE_URL   = "https://portal.distrinando.com.ar"
USERNAME   = "DEPORTE-CBA MAGRIS"   # el + en la URL equivale a espacio
PASSWORD   = "1973"
COMPANY_ID = "1"

# ─────────────────────────────────────────────────────────────────────────────


def generar_page_key():
    """Genera un PageKey aleatorio de 8 caracteres como usa el portal."""
    return ''.join(random.choices(string.ascii_letters + string.digits, k=8))


def log(text_widget, msg):
    text_widget.configure(state='normal')
    text_widget.insert('end', msg + '\n')
    text_widget.see('end')
    text_widget.configure(state='disabled')


def buscar_cliente(session, nombre, log_fn):
    """Busca un cliente por nombre y devuelve su CardCode."""
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
    log_fn(f"  Cliente encontrado: {cliente['CardName']} ({cliente['CardCode']})")
    return cliente["CardCode"], cliente["CardName"]


def obtener_datos_bp(session, card_code, log_fn):
    """Obtiene los datos completos del business partner."""
    r = session.post(
        f"{BASE_URL}/Sales/SalesOrder/GetBp",
        data=f"Id={card_code}&LocalCurrency=ARS",
        headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                 "X-Requested-With": "XMLHttpRequest"}
    )
    return r.json()


def buscar_item(session, sku, card_code, page_key, log_fn):
    """Paso 1: Busca el ítem via _Items (DataTable server-side). Devuelve el primer resultado o None."""
    payload = (
        "draw=1"
        "&columns[0][data]=ItemCode&columns[0][name]=ItemCode&columns[0][searchable]=true&columns[0][orderable]=false&columns[0][search][value]=&columns[0][search][regex]=false"
        "&columns[1][data]=ItemCode&columns[1][name]=ItemCode&columns[1][searchable]=true&columns[1][orderable]=true&columns[1][search][value]=&columns[1][search][regex]=false"
        "&columns[2][data]=ItemName&columns[2][name]=ItemName&columns[2][searchable]=true&columns[2][orderable]=true&columns[2][search][value]=&columns[2][search][regex]=false"
        "&columns[3][data]=&columns[3][name]=Stock&columns[3][searchable]=true&columns[3][orderable]=false&columns[3][search][value]=&columns[3][search][regex]=false"
        "&order[0][column]=0&order[0][dir]=asc&start=0&length=10&search[value]=&search[regex]=false"
        f"&pPageKey={page_key}"
        f"&pItemCode={requests.utils.quote(sku)}"
        "&pItemName="
        "&pCkCatalogueNum=false"
        f"&pCardCode={card_code}"
        "&pBPCatalogCode="
        "&pInventoryItem="
        "&pItemWithStock=N"
        "&pItemGroup=0"
    )
    r = session.post(
        f"{BASE_URL}/Sales/SalesOrder/_Items",
        data=payload,
        headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                 "X-Requested-With": "XMLHttpRequest"}
    )
    if r.status_code != 200:
        log_fn(f"    ⚠️  Error en búsqueda _Items: status {r.status_code}")
        return None
    data = r.json()
    resultados = data.get("data", [])
    if not resultados:
        log_fn(f"    ⚠️  No se encontró el ítem en la búsqueda")
        return None
    # Buscar coincidencia exacta, o tomar el primero
    for item in resultados:
        if item.get("ItemCode") == sku:
            return item
    log_fn(f"    ⚠️  No hay coincidencia exacta. Primer resultado: {resultados[0].get('ItemCode')}")
    return resultados[0]


def consultar_proyecto(session, card_code):
    """Consulta el proyecto una sola vez para el cliente (no depende del ítem)."""
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
    """Consulta el precio real via QueryManager."""
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
    """Agrega un ítem al formulario. Primero busca via _Items, luego carga via _ItemsForm."""
    from bs4 import BeautifulSoup
    from concurrent.futures import ThreadPoolExecutor

    # Paso 1: Buscar el ítem (esto inicializa el ítem en la sesión del servidor)
    item_search = buscar_item(session, sku, card_code, page_key, log_fn)
    if not item_search:
        return None
    item_code = item_search["ItemCode"]
    item_name = item_search.get("ItemName", item_code)
    log_fn(f"    → Encontrado: {item_name}")

    # Paso 2: Agregar al formulario via _ItemsForm + consultar precio en paralelo
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
        log_fn(f"    ⚠️  Error en _ItemsForm: status {r.status_code}")
        return None

    # Parsear el HTML para extraer datos
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
        log_fn(f"    → Precio: {price}")

        item_data = {
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
        return item_data
    except Exception as ex:
        log_fn(f"    ⚠️  Error parseando respuesta de _ItemsForm: {ex}")
        return None


def actualizar_cantidad(session, linea, cantidad, page_key, log_fn):
    """Actualiza la cantidad de una línea."""
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


def guardar_borrador(session, bp_data, lines, page_key, comentario, log_fn):
    """Guarda el pedido como borrador (SaveAsDraft=true) para revisión manual."""
    today = date.today()
    today = f"{today.month}/{today.day}/{today.year}"

    # Armar direcciones desde bp_data
    ship_addr = next((a for a in bp_data.get("Addresses", []) if a.get("AdresType") == "S"), {})
    bill_addr = next((a for a in bp_data.get("Addresses", []) if a.get("AdresType") == "B"), {})

    body = {
        "DocDate": today,
        "DocDueDate": today,
        "TaxDate": today,
        "CardCode": bp_data["CardCode"],
        "DocCur": "ARS",
        "DocRate": "1",
        "DocTotal": sum(l.get("Price", 0) * l.get("Quantity", 1) for l in lines),
        "CardName": bp_data["CardName"],
        "DiscPrcnt": "0",
        "CntctCode": str(bp_data.get("ListContact", [{}])[0].get("CntctCode", "")),
        "SlpCode": str(bp_data.get("SlpCode", "0")),
        "TrnspCode": None,
        "NumAtCard": "",
        "CancelDate": None,
        "ReqDate": None,
        "OwnerCode": "0",
        "Comments": comentario,
        "PageKey": page_key,
        "SOAddress": {
            "DocEntry": "0",
            "StreetS": ship_addr.get("Street", ""),
            "StreetB": bill_addr.get("Street", ""),
            "StreetNoS": "",
            "StreetNoB": "",
            "BlockS": "",
            "BlockB": "",
            "CityS": ship_addr.get("City", ""),
            "CityB": bill_addr.get("City", ""),
            "ZipCodeS": ship_addr.get("ZipCode", ""),
            "ZipCodeB": bill_addr.get("ZipCode", ""),
            "CountyS": "",
            "CountyB": "",
            "StateS": None,
            "StateB": None,
            "CountryS": None,
            "CountryB": None,
            "BuildingS": "",
            "BuildingB": "",
            "GlbLocNumS": "",
            "GlbLocNumB": ""
        },
        "ShipToCode": bp_data.get("ShipToDef", "ENTREGAR EN"),
        "PayToCode": bp_data.get("BillToDef", "FACTURAR A"),
        "GroupNum": str(bp_data.get("ListNum", "")),
        "PeyMethod": "",
        "ListItem": [],
        "Lines": lines,
        "ListFreight": [],
        "MappedUdf": [],
        "From": "MyDocuments",
        "UrlFrom": "/Sales/SalesDraft/Index",
        "QuickOrderId": "",
        "SaveAsDraft": "true"
    }

    first_line = json.dumps(lines[0]) if lines else "vacio"
    log_fn("\n🔍 Lines[0]: " + first_line)
    r = session.post(
        f"{BASE_URL}/Sales/SalesOrder/Add",
        json=body,
        headers={"Content-Type": "application/json; charset=UTF-8",
                 "X-Requested-With": "XMLHttpRequest"}
    )
    return r


def correr_carga(ruta_excel, nombre_cliente, descripcion, log_fn, btn_cargar):
    """Función principal que ejecuta toda la carga."""
    try:
        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
            "X-Requested-With": "XMLHttpRequest",
            "Accept-Language": "es-ES,es;q=0.9",
            "Origin": BASE_URL,
            "Referer": f"{BASE_URL}/Sales/SalesOrder/ActionPurchaseOrder?ActionPurchaseOrder=Add&IdPO=0&fromController=MyD",
        })

        # ── 1. LOGIN ─────────────────────────────────────────────────────────
        log_fn("🔐 Iniciando sesión...")
        r = session.post(
            f"{BASE_URL}/Login/Signin",
            data=f"username={requests.utils.quote(USERNAME)}&password={PASSWORD}&rememberMe=undefined",
            headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
        )
        if r.status_code != 200:
            log_fn(f"❌ Error en login: {r.status_code}")
            return

        r2 = session.post(
            f"{BASE_URL}/Login/SigninCompany",
            data=f"CompanyId={COMPANY_ID}",
            headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
        )
        log_fn("✅ Sesión iniciada correctamente")

        # ── Cargar la página del formulario para inicializar la sesión ────────
        FORM_URL = f"{BASE_URL}/Sales/SalesOrder/ActionPurchaseOrder?ActionPurchaseOrder=Add&IdPO=0&fromController=MyDocuments"
        log_fn("🌐 Cargando formulario de pedido...")
        from bs4 import BeautifulSoup
        form_r = session.get(FORM_URL, headers={"Referer": BASE_URL})
        # Setear el Referer para todos los requests siguientes
        session.headers.update({"Referer": FORM_URL})

        # Extraer el PageKey del formulario (input #Pagekey generado por el servidor)
        form_soup = BeautifulSoup(form_r.text, 'html.parser')
        pk_input = form_soup.find(id='Pagekey') or form_soup.find(id='PageKey') or form_soup.find(id='pageKey')
        if pk_input:
            page_key = pk_input.get('value', '')
        else:
            # Fallback: buscar en el HTML con regex
            import re
            pk_match = re.search(r'id=["\'](?:P|p)age[Kk]ey["\'][^>]*value=["\']([^"\']+)', form_r.text)
            if pk_match:
                page_key = pk_match.group(1)
            else:
                page_key = generar_page_key()
                log_fn("⚠️  No se encontró PageKey en el formulario, usando uno generado")

        # ── 2. BUSCAR CLIENTE ─────────────────────────────────────────────────
        log_fn(f"\n👤 Buscando cliente: {nombre_cliente}")
        card_code, card_name = buscar_cliente(session, nombre_cliente, log_fn)
        if not card_code:
            log_fn(f"❌ No se encontró ningún cliente con el nombre '{nombre_cliente}'")
            return

        # ── 3. OBTENER DATOS DEL BP ───────────────────────────────────────────
        log_fn(f"  Obteniendo datos del cliente...")
        bp_data = obtener_datos_bp(session, card_code, log_fn)

        # ── 4. INICIALIZAR FORMULARIO CON PAGE KEY ────────────────────────────
        log_fn(f"\n📋 PageKey del pedido: {page_key}")
        today = date.today()
        today_fmt = f"{today.month}/{today.day}/{today.year}"
        session.post(
            f"{BASE_URL}/Sales/SalesOrder/UpdateRateList",
            data=f"DocDate={requests.utils.quote(today_fmt)}&pPageKey={page_key}",
            headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
        )
        # Llamadas que el browser hace al cargar el formulario
        session.post(f"{BASE_URL}/Sales/SalesOrder/_GetDistributionRuleList",
                     headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
        session.post(f"{BASE_URL}/Sales/SalesOrder/_GetGLAccountList",
                     data=f"pPageKey={page_key}",
                     headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
        session.post(f"{BASE_URL}/Sales/SalesOrder/GetItemsModel",
                     headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})

        # ── 5. LEER EXCEL ─────────────────────────────────────────────────────
        log_fn(f"\n📊 Leyendo Excel...")
        wb = openpyxl.load_workbook(ruta_excel)
        ws = wb.active
        items_excel = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            sku = str(row[0]).strip() if row[0] else None
            cantidad = int(row[1]) if row[1] else 1
            if sku:
                items_excel.append((sku, cantidad))
        log_fn(f"  {len(items_excel)} ítems encontrados en el Excel")

        # ── 6. CONSULTAR PROYECTO (una sola vez para el cliente) ──────────────
        project = consultar_proyecto(session, card_code)
        if project:
            log_fn(f"  Proyecto: {project}")

        # ── 7. AGREGAR ÍTEMS ──────────────────────────────────────────────────
        log_fn(f"\n📦 Agregando ítems al pedido...")
        lines = []
        for idx, (sku, cantidad) in enumerate(items_excel):
            log_fn(f"  [{idx+1}/{len(items_excel)}] SKU: {sku} | Cantidad: {cantidad}")
            item_data = agregar_item(session, sku, card_code, page_key, project, log_fn)
            if item_data:
                item_data["Quantity"] = cantidad
                item_data["LineNum"] = str(idx)
                actualizar_cantidad(session, item_data, cantidad, page_key, log_fn)
                lines.append({
                    "ItemCode": item_data.get("ItemCode", sku),
                    "Dscription": item_data.get("ItemName", sku),
                    "Quantity": cantidad,
                    "Price": item_data.get("Price", 0),
                    "PriceBefDi": item_data.get("Price", 0),
                    "Currency": "ARS",
                    "LineNum": str(idx),
                    "WhsCode": item_data.get("WhsCode", "001"),
                    "OcrCode": item_data.get("OcrCode", ""),
                    "OcrCode2": item_data.get("OcrCode2", None),
                    "BaseType": "",
                    "BaseEntry": "",
                    "BaseLine": "",
                    "FreeTxt": "",
                    "GLAccount": {"FormatCode": ""},
                    "Project": item_data.get("Project", ""),
                    "UomCode": item_data.get("UomCode", ""),
                    "SerialBatch": "",
                    "ShipDate": None,
                    "MappedUdf": [],
                    "TaxCode": item_data.get("TaxCode", "IVA_21"),
                    "DiscPrcnt": str(item_data.get("DiscPrcnt", "0.0000")),
                    "VatPrcnt": str(item_data.get("VatPrcnt", "21.0000")),
                    "Freight": [{"ExpnsCode": "", "LineTotal": ""}]
                })
            else:
                log_fn(f"  ⚠️  No se encontró el SKU: {sku} — se omite")

        if not lines:
            log_fn("❌ No se pudo agregar ningún ítem. Verificá los SKUs.")
            return

        # ── 7. GUARDAR COMO BORRADOR ──────────────────────────────────────────
        log_fn(f"\n💾 Guardando borrador con {len(lines)} ítems...")
        r = guardar_borrador(session, bp_data, lines, page_key, descripcion, log_fn)

        if r.status_code == 200:
            try:
                resp_data = r.json()
                log_fn(f"\n✅ ¡Borrador guardado exitosamente!")
                log_fn(f"   Respuesta del servidor: {json.dumps(resp_data)[:200]}")
                log_fn(f"\n👉 Entrá al portal, buscá el borrador y confirmalo.")
            except Exception:
                log_fn(f"\n✅ Borrador guardado. Status: {r.status_code}")
        else:
            log_fn(f"\n❌ Error al guardar. Status: {r.status_code}")
            log_fn(f"   Respuesta: {r.text[:300]}")

    except Exception as ex:
        log_fn(f"\n❌ Error inesperado: {ex}")
        import traceback
        log_fn(traceback.format_exc())
    finally:
        btn_cargar.configure(state='normal')


# ─── INTERFAZ GRÁFICA ────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    root.title("Carga de Pedidos — OnePortal")
    root.geometry("620x540")
    root.resizable(False, False)
    root.configure(bg="#f0f0f0")

    # ── Frame principal ───────────────────────────────────────────────────────
    frame = ttk.Frame(root, padding=20)
    frame.pack(fill='both', expand=True)

    ttk.Label(frame, text="Carga de Pedidos al Portal", font=("Helvetica", 14, "bold")).grid(
        row=0, column=0, columnspan=3, pady=(0, 20))

    # ── Archivo Excel ─────────────────────────────────────────────────────────
    ttk.Label(frame, text="Archivo Excel (SKU / Cantidad):").grid(row=1, column=0, sticky='w')
    excel_var = tk.StringVar()
    ttk.Entry(frame, textvariable=excel_var, width=42).grid(row=1, column=1, padx=5)

    def browse():
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            excel_var.set(path)

    ttk.Button(frame, text="Buscar", command=browse).grid(row=1, column=2)

    # ── Nombre del cliente ────────────────────────────────────────────────────
    ttk.Label(frame, text="Nombre del cliente:").grid(row=2, column=0, sticky='w', pady=(12, 0))
    cliente_var = tk.StringVar()
    ttk.Entry(frame, textvariable=cliente_var, width=42).grid(row=2, column=1, padx=5, pady=(12, 0))

    # ── Descripción ───────────────────────────────────────────────────────────
    ttk.Label(frame, text="Descripción / Comentario:").grid(row=3, column=0, sticky='nw', pady=(12, 0))
    desc_text = tk.Text(frame, width=42, height=3, font=("TkDefaultFont", 10))
    desc_text.grid(row=3, column=1, padx=5, pady=(12, 0))

    # ── Log ───────────────────────────────────────────────────────────────────
    ttk.Label(frame, text="Log de ejecución:").grid(row=4, column=0, sticky='nw', pady=(16, 0))
    log_text = tk.Text(frame, width=60, height=12, state='disabled', font=("Courier", 9))
    log_text.grid(row=5, column=0, columnspan=3, pady=(4, 12))

    scrollbar = ttk.Scrollbar(frame, orient='vertical', command=log_text.yview)
    log_text.configure(yscrollcommand=scrollbar.set)

    # ── Botón cargar ──────────────────────────────────────────────────────────
    def iniciar_carga():
        ruta = excel_var.get().strip()
        cliente = cliente_var.get().strip()
        descripcion = desc_text.get("1.0", "end").strip()

        if not ruta:
            messagebox.showwarning("Falta dato", "Seleccioná un archivo Excel.")
            return
        if not cliente:
            messagebox.showwarning("Falta dato", "Ingresá el nombre del cliente.")
            return

        btn_cargar.configure(state='disabled')
        log_text.configure(state='normal')
        log_text.delete("1.0", "end")
        log_text.configure(state='disabled')

        threading.Thread(
            target=correr_carga,
            args=(ruta, cliente, descripcion, lambda m: log(log_text, m), btn_cargar),
            daemon=True
        ).start()

    btn_cargar = ttk.Button(frame, text="▶  Iniciar carga", command=iniciar_carga)
    btn_cargar.grid(row=6, column=0, columnspan=3, pady=(0, 8))

    root.mainloop()


if __name__ == "__main__":
    main()
