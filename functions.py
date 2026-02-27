import requests
import pandas as pd
import json
import datetime as dt
import time
import io
import threading
import pytz
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

def tp_tgt(username,password):
    """
    Creates a Ticket Granting Ticket (TGT) for accessing the EPIAS Transparency Platform API.

    Returns:
        str: The TGT string that is required for authenticated API calls.
    """
    url = "https://giris.epias.com.tr/cas/v1/tickets"
    data = {
        "username": username,
        "password": password,
    }

    headers = {
        "Accept": "text/plain", 
        "Content-Type": "application/x-www-form-urlencoded"
    }

    r = requests.post(url, data=data, headers=headers)
    return r.text

def transparency_call(tgt,method, service, endpoint, body):
    host = "https://seffaflik.epias.com.tr/"
    url = host + service + endpoint
    headers = {
        "Accept-Language": "en",
        "Accept": "application/json",
        "Content-Type": "application/json",
        "TGT": tgt
    }

    response = requests.request(method, url, headers=headers, json=body)

    return response


# ── GİP Transaction History ────────────────────────────────────────────────


def fetch_transaction_history_page(tgt, start_iso, end_iso, page_number, page_size=9999):
    """
    GİP işlem geçmişi endpoint'inden tek bir sayfa çeker.

    Parameters
    ----------
    tgt          : str  — tp_tgt() ile alınan TGT
    start_iso    : str  — "2025-12-01T00:00:00+03:00"
    end_iso      : str  — "2025-12-08T00:00:00+03:00"
    page_number  : int  — 1'den başlayan sayfa numarası
    page_size    : int  — Sayfa başı kayıt sayısı (max 9999)

    Returns
    -------
    dict — parse edilmiş JSON response

    Raises
    ------
    RuntimeError — HTTP 200 dışı yanıt
    """
    body = {
        "startDate": start_iso,
        "endDate": end_iso,
        "page": {
            "number": page_number,
            "size": page_size,
        }
    }
    response = transparency_call(
        tgt=tgt,
        method="POST",
        service="electricity-service/",
        endpoint="v1/markets/idm/data/transaction-history",
        body=body
    )
    if response.status_code == 401:
        raise RuntimeError("Oturum süresi doldu. Lütfen yeniden giriş yapın.")
    if response.status_code != 200:
        # JSON hata mesajını parse et, olmazsa ham metni göster
        try:
            err_data = response.json()
            errors = err_data.get("errors") or []
            if errors:
                msgs = " | ".join(
                    f"{e.get('errorCode', '')} {e.get('errorMessage', '')}"
                    for e in errors
                )
                raise RuntimeError(f"API hatası (HTTP {response.status_code}): {msgs}")
        except (ValueError, AttributeError):
            pass
        raise RuntimeError(
            f"API hatası (HTTP {response.status_code}): {response.text}"
        )
    return response.json()


def _parse_response_content(data):
    """
    API yanıtından items ve page bilgisini çıkarır.
    Hem body.content yapısını hem de düz yapıyı destekler.
    """
    try:
        content = data["body"]["content"]
        return content["items"], content["page"]
    except (KeyError, TypeError):
        return data["items"], data["page"]


def fetch_all_transaction_history(tgt, start_iso, end_iso, progress_callback=None, page_size=9999):
    """
    GİP işlem geçmişini sayfalandırarak tamamen çeker.

    Parameters
    ----------
    tgt               : str
    start_iso         : str  — "2025-12-01T00:00:00+03:00"
    end_iso           : str  — "2025-12-08T00:00:00+03:00"
    progress_callback : callable(current_page, total_pages) | None
    page_size         : int

    Returns
    -------
    pd.DataFrame — sütunlar: date, hour, contractName, price, quantity, id

    Raises
    ------
    RuntimeError — kimlik doğrulama hatası veya API hatası
    """
    first_page_data = fetch_transaction_history_page(
        tgt, start_iso, end_iso, page_number=1, page_size=page_size
    )
    items, page_info = _parse_response_content(first_page_data)

    total_records = page_info["total"]
    total_pages = max(1, (total_records + page_size - 1) // page_size)

    all_items = list(items)

    if progress_callback:
        progress_callback(1, total_pages)

    for page_num in range(2, total_pages + 1):
        page_data = fetch_transaction_history_page(
            tgt, start_iso, end_iso, page_number=page_num, page_size=page_size
        )
        page_items, _ = _parse_response_content(page_data)
        all_items.extend(page_items)

        if progress_callback:
            progress_callback(page_num, total_pages)

        time.sleep(0.05)

    return pd.DataFrame(all_items)


def fetch_all_transaction_history_chunked(
    tgt,
    start_iso,
    end_iso,
    progress_callback=None,
    slot_progress_callback=None,
    chunk_days=None,
    max_workers=4,
    max_retries=3,
    retry_base_delay=5.0,
):
    """
    Uzun tarih aralıklarını API'nin 7 günlük sınırını aşmadan paralel olarak çeker.

    Tarih aralığını `chunk_days` günlük parçalara böler ve `max_workers` kadar
    chunk'ı eş zamanlı olarak çeker.

    chunk_days otomatik hesaplama kuralı (chunk_days=None ise):
      - total_days <= max_workers * 3  →  chunk_days = 3
      - total_days >  max_workers * 3  →  chunk_days = min(ceil(total_days / max_workers), 7)
    Böylece her worker varsayılan olarak ~3 günlük veriden sorumlu tutulur;
    tarih aralığı genişledikçe yük eşit dağıtılır (API 7 gün limitine uyulur).

    Parameters
    ----------
    tgt                    : str   — TGT token
    start_iso              : str   — "2025-01-01T00:00:00+03:00"
    end_iso                : str   — "2025-12-31T23:59:59+03:00"
    progress_callback      : callable(pct: float 0-1, text: str) | None
                             Genel ilerleme; ana thread'den çağrılır.
    slot_progress_callback : callable(slot_idx: int, pct: float, text: str) | None
                             Worker başına ilerleme; ana thread'den çağrılır.
    chunk_days             : int | None — Chunk başı gün sayısı; None → yukarıdaki kurala göre otomatik hesaplanır
    max_workers            : int   — Eş zamanlı chunk sayısı
    max_retries            : int   — Geçici hatalarda yeniden deneme sayısı
    retry_base_delay       : float — Exponential backoff base süresi (saniye): 5s→10s→20s

    Returns
    -------
    pd.DataFrame — Tüm chunk'ların sıralı birleşik verisi

    Raises
    ------
    RuntimeError — Auth hatası veya max_retries aşıldıysa
    """
    from concurrent.futures import ThreadPoolExecutor, wait as _futures_wait, FIRST_COMPLETED
    import datetime as _dt

    import math as _math

    start_dt = _dt.datetime.fromisoformat(start_iso)
    end_dt = _dt.datetime.fromisoformat(end_iso)

    # chunk_days otomatik hesapla
    if chunk_days is None:
        total_days = (end_dt.date() - start_dt.date()).days + 1
        if total_days <= max_workers * 3:
            chunk_days = 3
        else:
            chunk_days = min(_math.ceil(total_days / max_workers), 7)

    # Chunk listesi oluştur
    chunks = []
    cs = start_dt
    while cs <= end_dt:
        ce = min(
            cs + _dt.timedelta(days=chunk_days - 1, hours=23, minutes=59, seconds=59),
            end_dt,
        )
        chunks.append((cs, ce))
        cs = cs + _dt.timedelta(days=chunk_days)

    total_chunks = len(chunks)

    # ── Per-slot durum takibi (thread-safe) ────────────────────────────────
    slot_states = {}        # slot_idx -> {current_page, total_pages, dates, done}
    state_lock = threading.Lock()
    thread_to_slot = {}     # thread ident -> slot_idx

    def _get_or_assign_slot():
        tid = threading.current_thread().ident
        with state_lock:
            if tid not in thread_to_slot:
                thread_to_slot[tid] = len(thread_to_slot)
            return thread_to_slot[tid]

    def _make_page_cb(slot_idx, chunk_dates):
        """Sayfa bazlı ilerlemeyi slot_states'e yazar (worker thread'den çağrılır)."""
        if slot_progress_callback is None:
            return None
        def page_cb(current_page, total_pages):
            with state_lock:
                slot_states[slot_idx] = {
                    "current_page": current_page,
                    "total_pages": total_pages,
                    "dates": chunk_dates,
                    "done": False,
                }
        return page_cb

    def _fetch_chunk(i, cs, ce):
        """Worker: tek chunk çeker, retry uygular."""
        slot_idx = _get_or_assign_slot()
        chunk_dates = f"{cs.strftime('%d.%m.%y')}–{ce.strftime('%d.%m.%y')}"
        page_cb = _make_page_cb(slot_idx, chunk_dates)
        for attempt in range(max_retries):
            try:
                df = fetch_all_transaction_history(tgt, cs.isoformat(), ce.isoformat(), page_cb)
                with state_lock:
                    slot_states[slot_idx] = {"dates": chunk_dates, "done": True}
                return i, df
            except RuntimeError as exc:
                if "Oturum süresi" in str(exc):
                    raise  # Auth hatası, retry işe yaramaz
                if attempt < max_retries - 1:
                    time.sleep(retry_base_delay * (2 ** attempt))  # 5s → 10s → 20s
                else:
                    raise

    def _flush_slot_ui():
        """slot_states'i okuyup slot_progress_callback'i çağırır. SADECE ana thread'den çağrılır."""
        if slot_progress_callback is None:
            return
        with state_lock:
            snapshot = dict(slot_states)
        for slot_idx, state in snapshot.items():
            if state.get("done"):
                slot_progress_callback(slot_idx, 1.0, f"✓ {state['dates']}")
            else:
                cp = state.get("current_page", 0)
                tp = state.get("total_pages", 1)
                pct = cp / max(tp, 1)
                slot_progress_callback(
                    slot_idx, pct,
                    f"Sayfa {cp}/{tp} — {state.get('dates', '')}",
                )

    # ── Paralel çekme ───────────────────────────────────────────────────────
    results = {}
    completed = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures_map = {
            executor.submit(_fetch_chunk, i, cs, ce): i
            for i, (cs, ce) in enumerate(chunks)
        }
        pending = set(futures_map.keys())

        while pending:
            # 300ms bekle — bu sürede worker'lar sayfa ilerlemelerini yazar
            done_set, pending = _futures_wait(pending, timeout=0.3, return_when=FIRST_COMPLETED)

            for future in done_set:
                i, df = future.result()  # Hata varsa burada ana thread'e fırlar
                results[i] = df
                completed += 1
                if progress_callback:
                    progress_callback(
                        completed / total_chunks,
                        f"{completed}/{total_chunks} chunk tamamlandı",
                    )

            # Slot barlarını ana thread'den güncelle
            _flush_slot_ui()

    all_frames = [results[i] for i in range(total_chunks)]
    return pd.concat(all_frames, ignore_index=True) if all_frames else pd.DataFrame()


# ── Kontrat Tarih/Saat Çözümleme ───────────────────────────────────────────


def parse_contract_datetimes(df):
    """
    Ham DataFrame'e kontrat ve işlem tarihlerini ekler.

    Kontrat adı formatı: PH + YY(2) + MM(2) + DD(2) + HH(2)
    Örnek: PH25120111 → teslimat 2025-12-01 11:00 Istanbul

    GİP kuralları:
      - T0 (açılış)  : teslimat günü - 1 gün, saat 18:00
      - T_close (kapanış): teslimat zamanı - 1 saat

    Eklenen sütunlar:
      transaction_dt  : tz-aware işlem zamanı (Europe/Istanbul)
      delivery_dt     : tz-aware teslimat zamanı
      t0_dt           : tz-aware kontrat açılış zamanı
      t_close_dt      : tz-aware kontrat kapanış zamanı
      delivery_date   : str "YYYY-MM-DD"
      delivery_hour   : int (0-23)
    """
    istanbul = pytz.timezone("Europe/Istanbul")
    df = df.copy()

    df["transaction_dt"] = pd.to_datetime(df["date"], utc=True).dt.tz_convert(istanbul)

    def _parse_delivery(name):
        yy = int(name[2:4])
        mm = int(name[4:6])
        dd = int(name[6:8])
        hh = int(name[8:10])
        naive = dt.datetime(2000 + yy, mm, dd, hh, 0, 0)
        return istanbul.localize(naive)

    df["delivery_dt"] = df["contractName"].apply(_parse_delivery)

    # T0: teslimat gününden 1 gün önce saat 18:00
    def _t0(delivery):
        prev_day = delivery.date() - dt.timedelta(days=1)
        naive = dt.datetime(prev_day.year, prev_day.month, prev_day.day, 18, 0, 0)
        return istanbul.localize(naive)

    df["t0_dt"] = df["delivery_dt"].apply(_t0)

    # T_close: teslimat zamanı - 1 saat
    df["t_close_dt"] = df["delivery_dt"] - pd.Timedelta(hours=1)

    df["delivery_date"] = df["delivery_dt"].dt.strftime("%Y-%m-%d")
    df["delivery_hour"] = df["delivery_dt"].dt.hour

    return df


# ── Filtre Uygulama ────────────────────────────────────────────────────────


def apply_filter_window(df, filter_def):
    """
    Tek bir filtre penceresini DataFrame'e uygular.

    filter_def şeması:
    {
        "name"     : str,           # görüntüleme adı
        "type"     : str,           # "first_n" | "last_n" | "custom"
        "n_min"    : int | None,    # first_n / last_n için dakika sayısı
        "start_min": int | None,    # custom için T0'dan başlangıç offseti (dk)
        "end_min"  : int | None,    # custom için T0'dan bitiş offseti (dk)
    }

    Pencere tanımları:
      first_n : [t0_dt, t0_dt + n_min]
      last_n  : [t_close_dt - n_min, t_close_dt]
      custom  : [t0_dt + start_min, t0_dt + end_min]

    Returns
    -------
    pd.DataFrame — filtre penceresine giren işlemler
    """
    df = df.copy()
    filter_type = filter_def["type"]

    if filter_type == "first_n":
        n = filter_def["n_min"]
        window_start = df["t0_dt"]
        window_end = df["t0_dt"] + pd.Timedelta(minutes=n)

    elif filter_type == "last_n":
        n = filter_def["n_min"]
        window_start = df["t_close_dt"] - pd.Timedelta(minutes=n)
        window_end = df["t_close_dt"]

    elif filter_type == "custom":
        start_m = filter_def["start_min"]
        end_m = filter_def["end_min"]
        window_start = df["t0_dt"] + pd.Timedelta(minutes=start_m)
        window_end = df["t0_dt"] + pd.Timedelta(minutes=end_m)

    else:
        raise ValueError(f"Bilinmeyen filtre tipi: {filter_type}")

    mask = (df["transaction_dt"] >= window_start) & (df["transaction_dt"] <= window_end)
    return df[mask]


# ── Ağırlıklı Ortalama Hesaplama ───────────────────────────────────────────


def calculate_weighted_averages(df, filter_defs):
    """
    Her kontrat için hacim ağırlıklı ortalama fiyatları hesaplar.

    Parameters
    ----------
    df          : pd.DataFrame — parse_contract_datetimes() çıktısı
    filter_defs : list[dict]   — filtre tanımları listesi

    Returns
    -------
    pd.DataFrame — kontrat başına bir satır, teslimat zamanına göre sıralı.

    Temel sütunlar:
      Kontrat Adı, Tarih, Saat, Kontrat Açılış, Kontrat Kapanış,
      Kontrat Süresi, Toplam İşlem Sayısı, Toplam Hacim, AOF

    Her filtre için ek sütunlar:
      {filtre_adı} İşlem Sayısı
      {filtre_adı} Hacim
      {filtre_adı} AOF

    Not: 1 lot = 0.1 MWh (GİP lot büyüklüğü)
    """
    LOT_TO_MWH = 0.1

    rows = []
    for contract, c_df in df.groupby("contractName"):
        delivery_dt = c_df["delivery_dt"].iloc[0]
        t0_dt = c_df["t0_dt"].iloc[0]
        t_close_dt = c_df["t_close_dt"].iloc[0]

        sure_dk = int((t_close_dt - t0_dt).total_seconds() / 60)

        total_q_all = c_df["quantity"].sum()
        if total_q_all > 0:
            aof_all = round(
                (c_df["price"] * c_df["quantity"]).sum() / total_q_all, 2
            )
        else:
            aof_all = None

        row = {
            "Kontrat Adı": contract,
            "Tarih": delivery_dt.strftime("%Y-%m-%d"),
            "Saat": delivery_dt.strftime("%H:%M"),
            "Kontrat Açılış": t0_dt.strftime("%Y-%m-%d %H:%M"),
            "Kontrat Kapanış": t_close_dt.strftime("%Y-%m-%d %H:%M"),
            "Kontrat Süresi": sure_dk,
            "Toplam İşlem Sayısı": len(c_df),
            "Toplam Hacim": round(total_q_all * LOT_TO_MWH, 3),
            "AOF": aof_all,
        }

        for fdef in filter_defs:
            fname = fdef["name"]
            filtered = apply_filter_window(c_df, fdef)

            total_q = filtered["quantity"].sum()
            if len(filtered) == 0 or total_q == 0:
                row[f"{fname} İşlem Sayısı"] = 0
                row[f"{fname} Hacim"] = 0.0
                row[f"{fname} AOF"] = None
            else:
                aof = (filtered["price"] * filtered["quantity"]).sum() / total_q
                row[f"{fname} İşlem Sayısı"] = len(filtered)
                row[f"{fname} Hacim"] = round(total_q * LOT_TO_MWH, 3)
                row[f"{fname} AOF"] = round(aof, 2)

        rows.append(row)

    result_df = pd.DataFrame(rows)
    result_df = result_df.sort_values(["Tarih", "Saat"]).reset_index(drop=True)
    return result_df


# ── Excel Dışa Aktarma ─────────────────────────────────────────────────────


def export_to_excel_bytes(result_df):
    """
    Sonuç DataFrame'ini biçimlendirilmiş Excel dosyası olarak döner.

    Returns
    -------
    bytes — .xlsx içeriği (st.download_button için)
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="GİP Filtreli Ortalama", index=False)
        ws = writer.sheets["GİP Filtreli Ortalama"]

        HEADER_BG = "2D4059"
        HEADER_FG = "FFFFFF"
        header_fill = PatternFill("solid", fgColor=HEADER_BG)
        header_font = Font(name="Calibri", size=10, bold=True, color=HEADER_FG)
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        data_font = Font(name="Calibri", size=10)
        align_center = Alignment(horizontal="center", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        thin = Side(style="thin", color="D9DEE4")
        data_border = Border(left=thin, right=thin, top=thin, bottom=thin)

        ws.row_dimensions[1].height = 45
        ws.freeze_panes = "D2"
        ws.sheet_view.zoomScale = 85

        BASE_COL_WIDTHS = {
            "Kontrat Adı": 14,
            "Tarih": 12,
            "Saat": 8,
            "Kontrat Açılış": 18,
            "Kontrat Kapanış": 18,
            "Kontrat Süresi": 12,
            "Toplam İşlem Sayısı": 14,
            "Toplam Hacim": 12,
            "AOF": 12,
        }

        for col_idx, col_name in enumerate(result_df.columns, start=1):
            col_letter = ws.cell(row=1, column=col_idx).column_letter

            h_cell = ws.cell(row=1, column=col_idx)
            h_cell.fill = header_fill
            h_cell.font = header_font
            h_cell.alignment = header_align

            width = BASE_COL_WIDTHS.get(col_name, 16)
            ws.column_dimensions[col_letter].width = width

        max_row = ws.max_row
        n_cols = len(result_df.columns)

        for row_idx in range(2, max_row + 1):
            for col_idx, col_name in enumerate(result_df.columns, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = data_font
                cell.border = data_border

                if "AOF" in col_name:
                    cell.alignment = align_right
                    cell.number_format = "#,##0.00"
                elif "Hacim" in col_name:
                    cell.alignment = align_right
                    cell.number_format = "#,##0.000"
                elif "Sayısı" in col_name or col_name == "Kontrat Süresi":
                    cell.alignment = align_center
                    cell.number_format = "#,##0"
                else:
                    cell.alignment = align_center

    output.seek(0)
    return output.getvalue()