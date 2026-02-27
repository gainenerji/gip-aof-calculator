import streamlit as st
import datetime as dt
import pytz

from functions import (
    tp_tgt,
    fetch_all_transaction_history_chunked,
    parse_contract_datetimes,
    calculate_weighted_averages,
    export_to_excel_bytes,
)

st.set_page_config(
    page_title="GİP AOF Hesaplayıcı",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Session state başlatma ─────────────────────────────────────────────────

if "tgt" not in st.session_state:
    st.session_state.tgt = None
if "auth_error" not in st.session_state:
    st.session_state.auth_error = None
if "filter_type" not in st.session_state:
    st.session_state.filter_type = "first_n"
if "filter_n" not in st.session_state:
    st.session_state.filter_n = 60
if "raw_df" not in st.session_state:
    st.session_state.raw_df = None
if "result_df" not in st.session_state:
    st.session_state.result_df = None
if "fetch_range" not in st.session_state:
    st.session_state.fetch_range = None
if "saved_username" not in st.session_state:
    st.session_state.saved_username = None
if "saved_password" not in st.session_state:
    st.session_state.saved_password = None
if "max_workers" not in st.session_state:
    st.session_state.max_workers = 4

# ── Yardımcı: aktif filtre adını üret ─────────────────────────────────────

def get_filter_label(filter_type, filter_n):
    if filter_type == "first_n":
        return f"İlk {filter_n} dakika"
    return f"Son {filter_n} dakika"

# ── Değişken başlangıç değerleri (sidebar içinde koşullu tanımlanıyor) ─────
fetch_btn = False
recalc_btn = False
start_date = None
end_date = None

# ── Sidebar ────────────────────────────────────────────────────────────────

with st.sidebar:

    # ── Kimlik Doğrulama ───────────────────────────────────────────────────
    if st.session_state.tgt is None:

        st.header("Kimlik Doğrulama")

        st.info(
            "Bu uygulamayı kullanmak için "
            "[EPIAS Şeffaflık Platformu](https://seffaflik.epias.com.tr) "
            "hesabınız gerekmektedir.",
            icon="ℹ️",
        )

        username = st.text_input(
            "E-posta",
            placeholder="kullanici@example.com",
            key="username_input",
        )
        password = st.text_input(
            "Şifre",
            type="password",
            key="password_input",
        )

        if st.button("Giriş Yap", use_container_width=True, type="primary"):
            if not username or not password:
                st.error("E-posta ve şifre gereklidir.")
            else:
                with st.spinner("Giriş yapılıyor..."):
                    try:
                        tgt_response = tp_tgt(username, password)
                        if "TGT" in tgt_response:
                            st.session_state.tgt = tgt_response.strip()
                            st.session_state.saved_username = username
                            st.session_state.saved_password = password
                            st.session_state.auth_error = None
                            st.rerun()
                        else:
                            st.session_state.auth_error = (
                                "Kimlik doğrulama başarısız. "
                                "Kullanıcı adı veya şifre hatalı."
                            )
                    except Exception as e:
                        st.session_state.auth_error = f"Bağlantı hatası: {e}"

        if st.session_state.auth_error:
            st.error(st.session_state.auth_error)

    else:
        # ── Giriş yapılmış: tarih + filtre + butonlar ──────────────────────

        st.success("Kimlik Doğrulandı")

        if st.button("TGT Yenile", use_container_width=True):
            if st.session_state.saved_username and st.session_state.saved_password:
                with st.spinner("TGT yenileniyor..."):
                    try:
                        tgt_response = tp_tgt(
                            st.session_state.saved_username,
                            st.session_state.saved_password,
                        )
                        if "TGT" in tgt_response:
                            st.session_state.tgt = tgt_response.strip()
                            st.toast("TGT başarıyla yenilendi.", icon="✅")
                        else:
                            st.error("TGT yenileme başarısız.")
                    except Exception as e:
                        st.error(f"Bağlantı hatası: {e}")
            else:
                st.error("Kimlik bilgileri bulunamadı. Lütfen yeniden giriş yapın.")

        st.divider()

        # ── Tarih Aralığı ──────────────────────────────────────────────────
        st.header("Tarih Aralığı")

        istanbul = pytz.timezone("Europe/Istanbul")
        today = dt.datetime.now(istanbul).date()
        tomorrow = today + dt.timedelta(days=1)

        start_date = st.date_input(
            "Başlangıç",
            value=today - dt.timedelta(days=7),
            max_value=tomorrow,
            key="start_date",
        )
        end_date = st.date_input(
            "Bitiş",
            value=today - dt.timedelta(days=1),
            max_value=tomorrow,
            key="end_date",
        )

        date_range_valid = start_date <= end_date
        if not date_range_valid:
            st.error("Başlangıç tarihi bitiş tarihinden sonra olamaz.")

        st.divider()

        # ── Filtre Penceresi ───────────────────────────────────────────────
        st.header("Filtre")

        tip_secim = st.selectbox(
            "Filtre tipi",
            options=["İlk N dakika", "Son N dakika"],
            index=0 if st.session_state.filter_type == "first_n" else 1,
            key="filter_type_select",
        )
        st.session_state.filter_type = "first_n" if tip_secim == "İlk N dakika" else "last_n"

        n_val = st.number_input(
            "N (dakika)",
            min_value=1,
            max_value=1440,
            value=st.session_state.filter_n,
            step=1,
            key="filter_n_input",
        )
        st.session_state.filter_n = int(n_val)

        filtre_adi = get_filter_label(st.session_state.filter_type, st.session_state.filter_n)
        st.caption(f"Aktif filtre: **{filtre_adi}**")

        st.divider()

        # ── Gelişmiş Ayarlar ───────────────────────────────────────────────
        st.header("Gelişmiş Ayarlar")

        st.session_state.max_workers = st.number_input(
            "Maksimum Paralel Worker Sayısı",
            min_value=1,
            max_value=10,
            value=st.session_state.max_workers,
            step=1,
            help=(
                "Eş zamanlı API isteği sayısı. "
                "Yüksek değer hızlıdır ancak rate limit riskini artırır.\n\n"
                "1 worker → ~40 dk/yıl\n"
                "4 worker → ~10 dk/yıl\n"
                "10 worker → ~4 dk/yıl"
            ),
            key="max_workers_input",
        )

        st.divider()

        # ── Aksiyon Butonları ──────────────────────────────────────────────
        fetch_btn = st.button(
            "Veri Çek ve Hesapla",
            use_container_width=True,
            type="primary",
            disabled=not date_range_valid,
        )

        recalc_btn = st.button(
            "Yeniden Hesapla",
            use_container_width=True,
            disabled=st.session_state.raw_df is None,
            help="Önbellekteki veriyle yeniden hesaplar — API'ye gitmez.",
        )

# ── Ana Alan ───────────────────────────────────────────────────────────────

st.title("GİP AOF Hesaplayıcı")
st.caption(
    "EPIAS Şeffaflık Platformu — Gün İçi Piyasası işlem geçmişinden "
    "kontrat bazlı hacim ağırlıklı ortalama fiyat hesaplaması."
)

# ── Veri çekme ─────────────────────────────────────────────────────────────

if st.session_state.tgt and fetch_btn:
    st.session_state.result_df = None

    # GİP kontratları teslimat gününden 1 gün önce 18:00'de açılır.
    # Örn: 1 Şubat kontratları → 31 Ocak 18:00'de açılır.
    # startDate'i 1 gün geri alıp 18:00'den başlatarak bu açılış işlemlerini dahil ediyoruz.
    # endDate'i seçilen günün 23:59:59'u olarak ayarlıyoruz;
    # API günü kapsayıcı yorumladığından +1 gün göndermek fazla kontrat getiriyor.
    start_iso = (start_date - dt.timedelta(days=1)).strftime("%Y-%m-%dT00:00:00+03:00")
    end_iso = end_date.strftime("%Y-%m-%dT23:59:59+03:00")

    _MAX_WORKERS = st.session_state.max_workers

    status = st.empty()
    prog = st.progress(0.0, text="Başlatılıyor...")

    # Her worker için ayrı bir progress bar
    worker_bars = [st.empty() for _ in range(_MAX_WORKERS)]

    def on_progress(pct_float, text):
        prog.progress(min(pct_float, 1.0), text=text)

    def on_slot_progress(slot_idx, pct, text):
        if 0 <= slot_idx < _MAX_WORKERS:
            worker_bars[slot_idx].progress(
                min(pct, 1.0),
                text=f"Worker {slot_idx + 1}: {text}",
            )

    def _clear_worker_bars():
        for bar in worker_bars:
            bar.empty()

    try:
        status.text("Veri çekiliyor...")
        raw_df = fetch_all_transaction_history_chunked(
            tgt=st.session_state.tgt,
            start_iso=start_iso,
            end_iso=end_iso,
            progress_callback=on_progress,
            slot_progress_callback=on_slot_progress,
            max_workers=_MAX_WORKERS,
        )

        _clear_worker_bars()
        prog.progress(1.0, text="Veri aktarımı tamamlandı.")
        status.text("İşlem zamanları çözümleniyor")
        parsed_df = parse_contract_datetimes(raw_df)

        # API geniş aralık döndürebilir; sadece seçilen teslimat günlerine ait kontratları tut
        from_str = start_date.strftime("%Y-%m-%d")
        to_str = end_date.strftime("%Y-%m-%d")
        parsed_df = parsed_df[
            (parsed_df["delivery_date"] >= from_str)
            & (parsed_df["delivery_date"] <= to_str)
        ].copy()

        filtre_adi = get_filter_label(st.session_state.filter_type, st.session_state.filter_n)
        filter_def = {
            "name": filtre_adi,
            "type": st.session_state.filter_type,
            "n_min": st.session_state.filter_n,
            "start_min": None,
            "end_min": None,
        }

        status.text("Ağırlıklı ortalamalar hesaplanıyor")
        result_df = calculate_weighted_averages(parsed_df, [filter_def])

        st.session_state.raw_df = parsed_df
        st.session_state.result_df = result_df
        st.session_state.fetch_range = (start_date, end_date)

        prog.empty()
        status.success(
            f"Tamamlandı — {len(raw_df):,} işlem, {len(result_df)} kontrat."
        )

    except RuntimeError as e:
        _clear_worker_bars()
        prog.empty()
        status.error(str(e))
    except Exception as e:
        _clear_worker_bars()
        prog.empty()
        status.error(f"Beklenmeyen hata: {e}")

# ── Yeniden hesaplama ──────────────────────────────────────────────────────

if st.session_state.tgt and recalc_btn and st.session_state.raw_df is not None:
    with st.spinner("Yeniden hesaplanıyor"):
        try:
            filtre_adi = get_filter_label(st.session_state.filter_type, st.session_state.filter_n)
            filter_def = {
                "name": filtre_adi,
                "type": st.session_state.filter_type,
                "n_min": st.session_state.filter_n,
                "start_min": None,
                "end_min": None,
            }
            result_df = calculate_weighted_averages(st.session_state.raw_df, [filter_def])
            st.session_state.result_df = result_df
            st.success(f"Yeniden hesaplama tamamlandı — filtre: {filtre_adi}")
        except Exception as e:
            st.error(f"Hesaplama hatası: {e}")

# ── Sonuçlar ───────────────────────────────────────────────────────────────

if st.session_state.result_df is not None:
    result_df = st.session_state.result_df
    fetch_start, fetch_end = st.session_state.fetch_range or (None, None)

    st.divider()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Kontrat Sayısı", f"{len(result_df)}")
    m2.metric("Toplam İşlem", f"{result_df['Toplam İşlem Sayısı'].sum():,}")
    m3.metric("Toplam Hacim (MWh)", f"{result_df['Toplam Hacim'].sum():,.1f}")
    if fetch_start and fetch_end:
        m4.metric(
            "Dönem",
            f"{fetch_start.strftime('%d.%m.%Y')} – {fetch_end.strftime('%d.%m.%Y')}",
        )

    st.divider()
    st.subheader("Sonuç Tablosu")

    # WA Fiyat ve Hacim sütunlarını formatla
    column_config = {}
    for col in result_df.columns:
        if "AOF" in col:
            column_config[col] = st.column_config.NumberColumn(format="%.2f")
        elif "Hacim" in col:
            column_config[col] = st.column_config.NumberColumn(format="%.3f")

    st.dataframe(
        result_df,
        use_container_width=True,
        hide_index=True,
        column_config=column_config,
    )

    st.divider()

    excel_bytes = export_to_excel_bytes(result_df)
    fname_start = fetch_start.strftime("%Y%m%d") if fetch_start else "start"
    fname_end = fetch_end.strftime("%Y%m%d") if fetch_end else "end"
    filename = f"GIP_WA_{fname_start}_{fname_end}.xlsx"

    st.download_button(
        label="Excel Olarak İndir",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
