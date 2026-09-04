import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="Haziran Servis Dashboard")

DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

def get_existing_data_path():
    for fname in os.listdir(DATA_DIR):
        if fname.startswith("son_veri."):
            return os.path.join(DATA_DIR, fname)
    return None

UPLOAD_PASSWORD = st.secrets.get("upload_password", "ik2026") if hasattr(st, "secrets") else "ik2026"

st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 1.0rem; }
[data-testid="stMetricLabel"] { font-size: 0.72rem; }
[data-testid="stMetricDelta"] { font-size: 0.65rem; }
[data-testid="stMetric"] { padding: 0.35rem 0.25rem; }
</style>
""", unsafe_allow_html=True)

COMPANIES = [
    'Aralık Sigorta', 'Ekim Turizm', 'Eylül Girişim',
    'Haziran Servis', 'Intercity Yatırım Holding', 'Mart Denizcilik'
]
MONTHS = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos']

# ----- FORMAT YARDIMCILARI -----
def format_tl(value):
    if value is None:
        return "0TL"
    try:
        formatted = f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{formatted}TL"
    except:
        return f"{value:.2f}TL".replace(".", ",")

def format_tl_no_decimal(value):
    if value is None:
        return "0TL"
    try:
        formatted = f"{value:,.0f}".replace(",", ".")
        return f"{formatted}TL"
    except:
        return f"{value:.0f}TL"

def format_percent(value):
    if value is None:
        return "%0"
    try:
        formatted = f"{value:.2f}".replace(".", ",")
        return f"%{formatted}"
    except:
        return f"%{value:.2f}".replace(".", ",")

def format_number(value, decimals=1):
    if value is None:
        return "0"
    try:
        formatted = f"{value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return formatted
    except:
        return str(value)

def calc_diff(current, previous):
    if previous is None:
        return None
    return round(current - previous, 4)

def format_delta_tl(diff):
    if diff is None:
        return None
    sign = "+" if diff >= 0 else "-"
    return f"{sign}{format_tl(abs(diff))}"

def format_delta_percent(diff):
    if diff is None:
        return None
    sign = "+" if diff >= 0 else "-"
    return f"{sign}{format_percent(abs(diff))}"

def format_delta_number(diff, decimals=0):
    if diff is None:
        return None
    sign = "+" if diff >= 0 else "-"
    return f"{sign}{format_number(abs(diff), decimals)}"

# ----- NORMALİZASYON -----
def normalize_company_name(name):
    if not isinstance(name, str):
        return name
    name = name.strip()
    replacements = {
        'EKIM TURIZM': 'Ekim Turizm',
        'HAZIRAN': 'Haziran Servis',
        'Holding': 'Intercity Yatırım Holding'
    }
    return replacements.get(name, name)

def clean_columns(df):
    df.columns = [str(col).strip() for col in df.columns]
    return df

def get_month_cols(df):
    month_cols = []
    for col in df.columns:
        col_clean = str(col).strip()
        for m in MONTHS:
            if col_clean.lower() == m.lower():
                month_cols.append(col)
                break
    return month_cols

def clean_numeric_df(df):
    return df.apply(pd.to_numeric, errors='coerce').fillna(0)

def read_company_month_sheet(uploaded_file, sheet_name, total_label, agg='sum', months=None):
    if months is None:
        months = MONTHS
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
    df = clean_columns(df)
    first_col = df.columns[0]
    df[first_col] = df[first_col].apply(normalize_company_name)
    df = df.set_index(first_col)
    month_cols = get_month_cols(df)

    is_total_row = df.index.astype(str).str.strip().str.upper() == total_label.strip().upper()
    total_row = df[is_total_row].iloc[0] if is_total_row.any() else None
    df_companies = df[~is_total_row].copy()
    for m in month_cols:
        df_companies[m] = pd.to_numeric(df_companies[m], errors='coerce').fillna(0)

    monthly_totals = {}
    for m in months:
        if m not in month_cols:
            monthly_totals[m] = 0
            continue
        sheet_val = total_row.get(m) if total_row is not None else None
        if sheet_val is not None and pd.notna(sheet_val):
            monthly_totals[m] = sheet_val
        else:
            series = df_companies[m]
            monthly_totals[m] = series.sum() if agg == 'sum' else series.mean()
    return df_companies, monthly_totals

def safe_read_son(uploaded_file, sheet, skip, n_months=None):
    if n_months is None:
        n_months = len(MONTHS)
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet, skiprows=skip, header=None, nrows=1)
        if df.shape[1] >= n_months + 1:
            return df.iloc[0, 1:n_months + 1].values
        else:
            return [0] * n_months
    except:
        return [0] * n_months

@st.cache_data(show_spinner=False)
def load_data(_uploaded_file, cache_key=None):
    uploaded_file = _uploaded_file

    # ----- 1. KÜMÜLATİF GENEL TURNOVER -----
    df_gt, gt_monthly_totals = read_company_month_sheet(uploaded_file, 'genel.turnover', 'Genel Toplam', agg='sum')

    # ----- 2. KÜMÜLATİF GÖNÜLLÜ TURNOVER -----
    df_gon, gon_monthly_totals = read_company_month_sheet(uploaded_file, 'gonullu.turnover', 'Genel Toplam', agg='sum')

    # ----- 3. AYLIK GENEL TURNOVER -----
    df_aylik_gt, aylik_gt_totals = read_company_month_sheet(uploaded_file, 'aylik.turnover', 'Genel Toplam', agg='sum')

    # ----- 4. AYLIK GÖNÜLLÜ TURNOVER -----
    df_aylik_gon, aylik_gon_totals = read_company_month_sheet(uploaded_file, 'aylik.gonullu.turnover', 'Genel Toplam', agg='sum')

    # ----- 5. RAPOR ORANI -----
    df_rapor = pd.read_excel(uploaded_file, sheet_name='rapor_oran', header=0)
    df_rapor = clean_columns(df_rapor)
    month_cols = get_month_cols(df_rapor)
    df_rapor['Şirket'] = df_rapor.iloc[:, 0].apply(normalize_company_name)
    df_rapor = clean_numeric_df(df_rapor.set_index('Şirket')[month_cols])

    # ----- 6. ÇALIŞAN SAYISI -----
    df_calisan = pd.read_excel(uploaded_file, sheet_name='calisan.sayisi', header=0)
    df_calisan = clean_columns(df_calisan)
    month_cols = get_month_cols(df_calisan)
    df_calisan['Şirket'] = df_calisan.iloc[:, 0].apply(normalize_company_name)
    df_calisan = clean_numeric_df(df_calisan.set_index('Şirket')[month_cols])

    # ----- 7. NET KÖK ÜCRET -----
    df_net = pd.read_excel(uploaded_file, sheet_name='kok.ucret', header=0)
    df_net = clean_columns(df_net)
    month_cols = get_month_cols(df_net)
    df_net['Şirket'] = df_net.iloc[:, 0].apply(normalize_company_name)
    df_net = clean_numeric_df(df_net.set_index('Şirket')[month_cols])

    # ----- 8. İŞVEREN MALİYETİ -----
    df_isv = pd.read_excel(uploaded_file, sheet_name='isveren.maliyet', header=0)
    df_isv = clean_columns(df_isv)
    month_cols = get_month_cols(df_isv)
    df_isv['Şirket'] = df_isv.iloc[:, 0].apply(normalize_company_name)
    df_isv = clean_numeric_df(df_isv.set_index('Şirket')[month_cols])

    # ----- 9. FM SAAT -----
    df_fm_saat = pd.read_excel(uploaded_file, sheet_name='fm.saat', header=0)
    df_fm_saat = clean_columns(df_fm_saat)
    month_cols = get_month_cols(df_fm_saat)
    df_fm_saat['Şirket'] = df_fm_saat.iloc[:, 0].apply(normalize_company_name)
    df_fm_saat = clean_numeric_df(df_fm_saat.set_index('Şirket')[month_cols])

    # ----- 10. FM TL MALİYET -----
    df_fm_tl = pd.read_excel(uploaded_file, sheet_name='fm.maliyet', header=0)
    df_fm_tl = clean_columns(df_fm_tl)
    month_cols = get_month_cols(df_fm_tl)
    df_fm_tl['Şirket'] = df_fm_tl.iloc[:, 0].apply(normalize_company_name)
    df_fm_tl = clean_numeric_df(df_fm_tl.set_index('Şirket')[month_cols])

    # ----- 11. İZİN GÜN -----
    df_izin_gun = pd.read_excel(uploaded_file, sheet_name='izin_gun', header=0)
    df_izin_gun = clean_columns(df_izin_gun)
    month_cols = get_month_cols(df_izin_gun)
    df_izin_gun['Şirket'] = df_izin_gun.iloc[:, 0].apply(normalize_company_name)
    df_izin_gun = clean_numeric_df(df_izin_gun.set_index('Şirket')[month_cols])

    # ----- 12. İZİN ÜCRET -----
    df_izin_ucret = pd.read_excel(uploaded_file, sheet_name='izin_ucret', header=0)
    df_izin_ucret = clean_columns(df_izin_ucret)
    month_cols = get_month_cols(df_izin_ucret)
    df_izin_ucret['Şirket'] = df_izin_ucret.iloc[:, 0].apply(normalize_company_name)
    df_izin_ucret = clean_numeric_df(df_izin_ucret.set_index('Şirket')[month_cols])

    # ----- 13. KIDEM TAZMİNATI -----
    df_kidem, kidem_totals = read_company_month_sheet(uploaded_file, 'kidem.tazminati', 'TOPLAM', agg='sum')

    # ----- 14. İHBAR TAZMİNATI -----
    df_ihbar, ihbar_totals = read_company_month_sheet(uploaded_file, 'ihbar.tazminati', 'TOPLAM', agg='sum')

    # ----- 15. KİŞİ BAŞI ORTALAMA MAAŞ (aylık) - kisi.basi.ort sayfasındaki "Genel Kişi Başı Ortalama Maaş" satırı -----
    df_kisi_basi, kisi_basi_genel = read_company_month_sheet(
        uploaded_file, 'kisi.basi.ort', 'Genel Kişi Başı Ortalama Maaş', agg='mean'
    )
    # Yıllık Ortalama sütununu bul (şirket bazlı yıllık ortalama için)
    yillik_ort_col = None
    for col in df_kisi_basi.columns:
        if str(col).strip().lower() in ('yıllık ortalama', 'yillik ortalama'):
            yillik_ort_col = col
            break

    # ----- 16. gnl.kisi.basi.ort sayfasından Genel Kişi Başı Ortalama Maaş (kümülatif ortalama) -----
    try:
        df_gnl = pd.read_excel(uploaded_file, sheet_name='gnl.kisi.basi.ort', header=0)
        df_gnl = clean_columns(df_gnl)
        if len(df_gnl) > 0:
            genel_satir = df_gnl.iloc[[0]]
            gnl_kisi_basi_vals = {}
            for m in MONTHS:
                if m in genel_satir.columns:
                    val = pd.to_numeric(genel_satir.iloc[0][m], errors='coerce')
                    gnl_kisi_basi_vals[m] = val if pd.notna(val) else 0
                else:
                    gnl_kisi_basi_vals[m] = 0
        else:
            gnl_kisi_basi_vals = {m: 0 for m in MONTHS}
    except Exception:
        gnl_kisi_basi_vals = {m: 0 for m in MONTHS}

    # ----- 17. KADIN ORANI -----
    df_kadin, kadin_genel = read_company_month_sheet(uploaded_file, 'kadin.erkek', 'Genel Kadın Oranı', agg='mean')

    # ----- 18. İLK 6 AY AYRILMA ORANI -----
    df_ilk6ay, ilk6ay_ortalama = read_company_month_sheet(uploaded_file, 'ilk.6ay', '__YOK__', agg='mean')

    # ----- 19. AY İÇİ GİRİŞ -----
    df_giris, giris_toplam = read_company_month_sheet(uploaded_file, 'aylik.giris', '__YOK__', agg='sum')

    # ----- 20. AY İÇİ ÇIKIŞ -----
    df_cikis, cikis_toplam = read_company_month_sheet(uploaded_file, 'aylik.cikis', '__YOK__', agg='sum')

    # ----- TOPLAM SATIRLARI (rapor_oran, calisan.sayisi, izin_gun, izin_ucret) -----
    genel_rapor = safe_read_son(uploaded_file, 'rapor_oran', 7)
    toplam_calisan = safe_read_son(uploaded_file, 'calisan.sayisi', 7)
    toplam_izin_gun = safe_read_son(uploaded_file, 'izin_gun', 7)
    toplam_izin_ucret = safe_read_son(uploaded_file, 'izin_ucret', 7)

    # ----- FM YAPAN LİSTESİ -----
    df_fm_yapan = pd.read_excel(uploaded_file, sheet_name='aylik.fm.yapan', header=0)
    df_fm_yapan = clean_columns(df_fm_yapan)
    if 'Şirket' in df_fm_yapan.columns:
        df_fm_yapan['Şirket'] = df_fm_yapan['Şirket'].apply(normalize_company_name)
    for m in get_month_cols(df_fm_yapan):
        df_fm_yapan[m] = pd.to_numeric(df_fm_yapan[m], errors='coerce').fillna(0)

    # ----- 21. LOKASYON BAZINDA ÇALIŞAN SAYISI -----
    df_lokasyon_cal = pd.read_excel(uploaded_file, sheet_name='lokasyon.bazinda.cal', header=0)
    df_lokasyon_cal = clean_columns(df_lokasyon_cal)
    first_col = df_lokasyon_cal.columns[0]
    df_lokasyon_cal = df_lokasyon_cal.set_index(first_col)
    month_cols = get_month_cols(df_lokasyon_cal)
    df_lokasyon_cal = clean_numeric_df(df_lokasyon_cal[month_cols])

    # ----- 22. LOKASYON BAZINDA KÖK ÜCRET -----
    df_lokasyon_ucret = pd.read_excel(uploaded_file, sheet_name='lokasyon.baz.kok.ucret', header=0)
    df_lokasyon_ucret = clean_columns(df_lokasyon_ucret)
    first_col = df_lokasyon_ucret.columns[0]
    df_lokasyon_ucret = df_lokasyon_ucret.set_index(first_col)
    month_cols = get_month_cols(df_lokasyon_ucret)
    df_lokasyon_ucret = clean_numeric_df(df_lokasyon_ucret[month_cols])

    # ----- VERİYİ BİRLEŞTİR -----
    data = {}
    for idx, m in enumerate(MONTHS):
        comp_data = {}
        for comp in COMPANIES:
            if comp in df_calisan.index:
                comp_data[comp] = {
                    'employees': df_calisan.loc[comp, m],
                    'devamsizlik': df_rapor.loc[comp, m] * 100,
                    'turnoverKumulatif': df_gt.loc[comp, m] * 100 if comp in df_gt.index else 0,
                    'turnoverGonulluKumulatif': df_gon.loc[comp, m] * 100 if comp in df_gon.index else 0,
                    'turnoverAylik': df_aylik_gt.loc[comp, m] * 100 if comp in df_aylik_gt.index else 0,
                    'turnoverGonulluAylik': df_aylik_gon.loc[comp, m] * 100 if comp in df_aylik_gon.index else 0,
                    'netKokUcret': df_net.loc[comp, m],
                    'isverenMaliyet': df_isv.loc[comp, m],
                    'fmSaat': df_fm_saat.loc[comp, m],
                    'fmTlMaliyet': df_fm_tl.loc[comp, m],
                    'izinGun': df_izin_gun.loc[comp, m],
                    'izinUcret': df_izin_ucret.loc[comp, m],
                    'kisiBasiOrt': df_kisi_basi.loc[comp, m] if comp in df_kisi_basi.index and m in df_kisi_basi.columns else 0,
                    'yillikOrtMaas': (df_kisi_basi.loc[comp, yillik_ort_col]
                                      if yillik_ort_col is not None and comp in df_kisi_basi.index else 0),
                    'kadinOrani': df_kadin.loc[comp, m] * 100 if comp in df_kadin.index and m in df_kadin.columns else 0,
                    'ilk6ayOrani': df_ilk6ay.loc[comp, m] * 100 if comp in df_ilk6ay.index and m in df_ilk6ay.columns else 0,
                    'aySekiceGiris': df_giris.loc[comp, m] if comp in df_giris.index and m in df_giris.columns else 0,
                    'aySekiceCikis': df_cikis.loc[comp, m] if comp in df_cikis.index and m in df_cikis.columns else 0,
                    'kidemTazminati': df_kidem.loc[comp, m] if comp in df_kidem.index and m in df_kidem.columns else 0,
                    'ihbarTazminati': df_ihbar.loc[comp, m] if comp in df_ihbar.index and m in df_ihbar.columns else 0,
                }
            else:
                comp_data[comp] = {key: 0 for key in ['employees', 'devamsizlik', 'turnoverKumulatif', 'turnoverGonulluKumulatif',
                                                      'turnoverAylik', 'turnoverGonulluAylik', 'netKokUcret',
                                                      'isverenMaliyet', 'fmSaat', 'fmTlMaliyet', 'izinGun', 'izinUcret',
                                                      'kisiBasiOrt', 'yillikOrtMaas', 'kadinOrani', 'ilk6ayOrani',
                                                      'aySekiceGiris', 'aySekiceCikis', 'kidemTazminati', 'ihbarTazminati']}

        # Genel toplamlar (sayfa son satırları)
        sheet_calisan = toplam_calisan[idx] if idx < len(toplam_calisan) else 0
        sheet_rapor = genel_rapor[idx] * 100 if idx < len(genel_rapor) else 0
        sheet_izin_gun = toplam_izin_gun[idx] if idx < len(toplam_izin_gun) else 0
        sheet_izin_ucret = toplam_izin_ucret[idx] if idx < len(toplam_izin_ucret) else 0

        calc_calisan = sum(v['employees'] for v in comp_data.values())
        calc_rapor = sum(v['devamsizlik'] for v in comp_data.values()) / len(COMPANIES)
        calc_izin_gun = sum(v['izinGun'] for v in comp_data.values())
        calc_izin_ucret = sum(v['izinUcret'] for v in comp_data.values())

        # Turnover genel toplamlar (kümülatif ve aylık)
        genel_kumulatif_turnover = gt_monthly_totals.get(m, 0) * 100
        genel_kumulatif_gonullu = gon_monthly_totals.get(m, 0) * 100
        genel_aylik_turnover = aylik_gt_totals.get(m, 0) * 100
        genel_aylik_gonullu = aylik_gon_totals.get(m, 0) * 100

        data[m] = {
            'companies': comp_data,
            'genelRaporOran': sheet_rapor if sheet_rapor else calc_rapor,
            'toplamCalisan': sheet_calisan if sheet_calisan else calc_calisan,
            'toplamIzinGun': sheet_izin_gun if sheet_izin_gun else calc_izin_gun,
            'toplamIzinUcret': sheet_izin_ucret if sheet_izin_ucret else calc_izin_ucret,
            'kidemTazminati': kidem_totals.get(m, 0),
            'ihbarTazminati': ihbar_totals.get(m, 0),
            'kisiBasiOrtGenel': kisi_basi_genel.get(m, 0),
            'kisiBasiOrtGenelKumulatif': gnl_kisi_basi_vals.get(m, 0),
            'kadinOraniGenel': kadin_genel.get(m, 0) * 100,
            'ilk6ayOraniGenel': ilk6ay_ortalama.get(m, 0) * 100,
            'girisToplam': giris_toplam.get(m, 0),
            'cikisToplam': cikis_toplam.get(m, 0),
            'genelKumulatifTurnover': genel_kumulatif_turnover,
            'genelKumulatifGonullu': genel_kumulatif_gonullu,
            'genelAylikTurnover': genel_aylik_turnover,
            'genelAylikGonullu': genel_aylik_gonullu,
        }

    # Şirket bazlı toplam turnover (kümülatif son sütun)
    df_gt_total = pd.read_excel(uploaded_file, sheet_name='genel.turnover', header=0)
    df_gt_total = clean_columns(df_gt_total)
    total_col = None
    for col in df_gt_total.columns:
        if 'toplam' in str(col).lower():
            total_col = col
            break
    if total_col is None:
        total_col = df_gt_total.columns[-1]
    df_gt_total['Şirket'] = df_gt_total.iloc[:, 0].apply(normalize_company_name)
    df_gt_total = df_gt_total.set_index('Şirket')
    turnover_sirket_toplam = df_gt_total[~df_gt_total.index.isna()][total_col] * 100

    df_gon_total = pd.read_excel(uploaded_file, sheet_name='gonullu.turnover', header=0)
    df_gon_total = clean_columns(df_gon_total)
    total_col_gon = None
    for col in df_gon_total.columns:
        if 'toplam' in str(col).lower():
            total_col_gon = col
            break
    if total_col_gon is None:
        total_col_gon = df_gon_total.columns[-1]
    df_gon_total['Şirket'] = df_gon_total.iloc[:, 0].apply(normalize_company_name)
    df_gon_total = df_gon_total.set_index('Şirket')
    turnover_sirket_gonullu = df_gon_total[~df_gon_total.index.isna()][total_col_gon] * 100

    turnover_sirket_bazli = pd.DataFrame({
        'Toplam': [turnover_sirket_toplam.get(c, 0) for c in COMPANIES],
        'Gönüllü': [turnover_sirket_gonullu.get(c, 0) for c in COMPANIES]
    }, index=COMPANIES)

    return {
        'by_month': data,
        'fm_yapan': df_fm_yapan,
        'turnoverSirketBazli': turnover_sirket_bazli,
        'lokasyon_cal': df_lokasyon_cal,
        'lokasyon_ucret': df_lokasyon_ucret,
    }


# ----- ANA UYGULAMA (SADECE HAZİRAN SERVİS) -----
def main():
    st.title("📊 Haziran Servis Özel Dashboard")

    existing_path = get_existing_data_path()
    with st.sidebar:
        with st.expander("🔒 Veriyi Güncelle", expanded=not existing_path):
            pwd = st.text_input("Şifre", type="password")
            new_file = st.file_uploader("Yeni Excel dosyası (.xlsx / .xlsb)", type=["xlsx", "xlsb"], key="haziran_admin_uploader")
            if new_file is not None:
                if pwd == UPLOAD_PASSWORD:
                    for fname in os.listdir(DATA_DIR):
                        if fname.startswith("son_veri."):
                            os.remove(os.path.join(DATA_DIR, fname))
                    ext = os.path.splitext(new_file.name)[1].lower()
                    new_path = os.path.join(DATA_DIR, f"son_veri{ext}")
                    with open(new_path, "wb") as f:
                        f.write(new_file.getbuffer())
                    st.cache_data.clear()
                    st.success("✅ Veri güncellendi.")
                    st.rerun()
                else:
                    st.error("Şifre yanlış.")

    st.markdown("---")

    data_path = get_existing_data_path()
    if not data_path:
        st.info("Henüz veri yüklenmedi. Soldaki '🔒 Veriyi Güncelle' panelinden bir Excel dosyası yükleyin.")
        return

    try:
        cache_key = f"{data_path}:{os.path.getmtime(data_path)}"
        all_data = load_data(data_path, cache_key=cache_key)
    except Exception as e:
        st.error(f"Hata oluştu: {e}")
        st.exception(e)
        st.stop()

    data = all_data['by_month']
    fm_yapan_df = all_data['fm_yapan']
    lokasyon_cal = all_data['lokasyon_cal']
    lokasyon_ucret = all_data['lokasyon_ucret']
    COMPANY = 'Haziran Servis'

    selected_month = st.selectbox("📅 Ay Seçin", MONTHS, index=len(MONTHS) - 1)
    month_idx = MONTHS.index(selected_month)
    prev_month = MONTHS[month_idx - 1] if month_idx > 0 else None

    hz = data[selected_month]['companies'][COMPANY]
    hz_prev = data[prev_month]['companies'][COMPANY] if prev_month else None

    def d(key):
        return hz.get(key, 0)

    def diff_of(key):
        if hz_prev is None:
            return None
        return calc_diff(hz.get(key, 0), hz_prev.get(key, 0))

    # ----- KPI KARTLARI (3 SATIR) -----
    # 1. satır: 7 kart
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
    col1.metric("👥 Çalışan Sayısı", format_number(d('employees'), 0), delta=format_delta_number(diff_of('employees'), 0))
    col2.metric("👩 Kadın Oranı", format_percent(d('kadinOrani')), delta=format_delta_percent(diff_of('kadinOrani')))
    col3.metric("📊 Raporlu Oran", format_percent(d('devamsizlik')), delta=format_delta_percent(diff_of('devamsizlik')))
    col4.metric("💼 İşveren Maliyeti", format_tl(d('isverenMaliyet')), delta=format_delta_tl(diff_of('isverenMaliyet')))
    col5.metric("💰 Net Kök Ücret", format_tl(d('netKokUcret')), delta=format_delta_tl(diff_of('netKokUcret')))
    col6.metric("⏱️ FM Saat", format_number(d('fmSaat'), 1), delta=format_delta_number(diff_of('fmSaat'), 1))
    col7.metric("💸 FM (Net TL)", format_tl(d('fmTlMaliyet')), delta=format_delta_tl(diff_of('fmTlMaliyet')))

    # 2. satır: 6 kart
    col8, col9, col10, col11, col12, col13 = st.columns(6)
    col8.metric("📅 İzin Gün Bakiyesi", format_number(d('izinGun'), 1), delta=format_delta_number(diff_of('izinGun'), 1))
    col9.metric("💎 İzin Ücreti (Net TL)", format_tl(d('izinUcret')), delta=format_delta_tl(diff_of('izinUcret')))
    col10.metric("🏷️ Kıdem Tazminatı (Net TL)", format_tl(d('kidemTazminati')), delta=format_delta_tl(diff_of('kidemTazminati')))
    col11.metric("📨 İhbar Tazminatı (Net TL)", format_tl(d('ihbarTazminati')), delta=format_delta_tl(diff_of('ihbarTazminati')))
    col12.metric("🧮 Kişi Başı Ort. Maaş (Net TL)", format_tl(d('kisiBasiOrt')), delta=format_delta_tl(diff_of('kisiBasiOrt')))
    col13.metric("🔄 Küm. Genel Turnover", format_percent(d('turnoverKumulatif')), delta=format_delta_percent(diff_of('turnoverKumulatif')))

    # 3. satır: 6 kart
    col14, col15, col16, col17, col18, col19 = st.columns(6)
    col14.metric("🚪 Küm. Gönüllü Turnover", format_percent(d('turnoverGonulluKumulatif')), delta=format_delta_percent(diff_of('turnoverGonulluKumulatif')))
    col15.metric("📈 Aylık Genel Turnover", format_percent(d('turnoverAylik')), delta=format_delta_percent(diff_of('turnoverAylik')))
    col16.metric("📉 Aylık Gönüllü Turnover", format_percent(d('turnoverGonulluAylik')), delta=format_delta_percent(diff_of('turnoverGonulluAylik')))
    col17.metric("⬆️ Ay İçi İşe Giren", format_number(d('aySekiceGiris'), 0), delta=format_delta_number(diff_of('aySekiceGiris'), 0))
    col18.metric("⬇️ Ay İçi İşten Ayrılan", format_number(d('aySekiceCikis'), 0), delta=format_delta_number(diff_of('aySekiceCikis'), 0))
    col19.metric("⏳ İlk 6 Ay Ayrılma Oranı", format_percent(d('ilk6ayOrani')), delta=format_delta_percent(diff_of('ilk6ayOrani')))

    st.markdown("---")

    # ----- KÜMÜLATİF TURNOVER TRENDİ -----
    st.subheader("📈 Haziran Servis - Kümülatif Turnover Trendi")
    trend_df = pd.DataFrame({
        'Ay': MONTHS,
        'Kümülatif Genel Turnover': [data[m]['companies'][COMPANY]['turnoverKumulatif'] for m in MONTHS],
        'Kümülatif Gönüllü Turnover': [data[m]['companies'][COMPANY]['turnoverGonulluKumulatif'] for m in MONTHS]
    })
    fig_trend = go.Figure()
    fig_trend.add_trace(go.Scatter(
        x=trend_df['Ay'], y=trend_df['Kümülatif Genel Turnover'],
        mode='lines+markers+text', name='Kümülatif Genel Turnover',
        line=dict(color='#f59e0b'), marker=dict(size=8),
        text=trend_df['Kümülatif Genel Turnover'].apply(format_percent),
        textposition='top center',
        hovertemplate='<b>%{x}</b><br>%{fullData.name}: %{text}<extra></extra>'
    ))
    fig_trend.add_trace(go.Scatter(
        x=trend_df['Ay'], y=trend_df['Kümülatif Gönüllü Turnover'],
        mode='lines+markers+text', name='Kümülatif Gönüllü Turnover',
        line=dict(color='#ec4899'), marker=dict(size=8),
        text=trend_df['Kümülatif Gönüllü Turnover'].apply(format_percent),
        textposition='bottom center',
        hovertemplate='<b>%{x}</b><br>%{fullData.name}: %{text}<extra></extra>'
    ))
    fig_trend.update_layout(
        height=400,
        margin=dict(l=10, r=10, t=30, b=10),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        yaxis_title='Turnover Oranı (%)'
    )
    st.plotly_chart(fig_trend, use_container_width=True)

    st.markdown("---")

    # ----- AYLIK TURNOVER GRAFİĞİ -----
    st.subheader("📊 Haziran Servis - Aylık Turnover")
    aylik_df = pd.DataFrame({
        'Ay': MONTHS,
        'Aylık Genel': [data[m]['companies'][COMPANY]['turnoverAylik'] for m in MONTHS],
        'Aylık Gönüllü': [data[m]['companies'][COMPANY]['turnoverGonulluAylik'] for m in MONTHS]
    })
    fig_aylik = go.Figure()
    fig_aylik.add_trace(go.Bar(x=aylik_df['Ay'], y=aylik_df['Aylık Genel'], name='Aylık Genel',
                                marker_color='#3b82f6', text=aylik_df['Aylık Genel'].apply(format_percent),
                                textposition='outside'))
    fig_aylik.add_trace(go.Bar(x=aylik_df['Ay'], y=aylik_df['Aylık Gönüllü'], name='Aylık Gönüllü',
                                marker_color='#8b5cf6', text=aylik_df['Aylık Gönüllü'].apply(format_percent),
                                textposition='outside'))
    fig_aylik.update_layout(barmode='group', height=350, margin=dict(l=10, r=10, t=30, b=10),
                             legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
    st.plotly_chart(fig_aylik, use_container_width=True)

    st.markdown("---")

    # ----- TÜM METRİKLER - AYLIK TREND (Ocak-Ağustos) -----
    st.subheader("📊 Haziran Servis - Tüm Metrikler (Ocak–Ağustos Trendi)")

    def series(key):
        return [data[m]['companies'][COMPANY][key] for m in MONTHS]

    def line_fig(title, series_dict, colors, fmt='num', decimals=1, yaxis_title=''):
        fig = go.Figure()
        for i, (name, vals) in enumerate(series_dict.items()):
            if fmt == 'percent':
                text = [format_percent(v) for v in vals]
            elif fmt == 'tl':
                text = [format_tl(v) for v in vals]
            else:
                text = [format_number(v, decimals) for v in vals]
            fig.add_trace(go.Scatter(
                x=MONTHS, y=vals, mode='lines+markers+text', name=name,
                line=dict(color=colors[i], width=3), marker=dict(size=7),
                text=text, textposition='top center' if i == 0 else 'bottom center',
                hovertemplate='<b>%{x}</b><br>' + name + ': %{text}<extra></extra>'
            ))
        fig.update_layout(title=title, height=340, margin=dict(l=10, r=10, t=40, b=10),
                           legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                           yaxis_title=yaxis_title)
        return fig

    def bar_fig(title, series_dict, colors, fmt='num', decimals=1, barmode='group'):
        fig = go.Figure()
        for i, (name, vals) in enumerate(series_dict.items()):
            if fmt == 'percent':
                text = [format_percent(v) for v in vals]
            elif fmt == 'tl':
                text = [format_tl(v) for v in vals]
            else:
                text = [format_number(v, decimals) for v in vals]
            fig.add_trace(go.Bar(x=MONTHS, y=vals, name=name, marker_color=colors[i],
                                  text=text, textposition='outside',
                                  hovertemplate='<b>%{x}</b><br>' + name + ': %{text}<extra></extra>'))
        fig.update_layout(title=title, height=340, margin=dict(l=10, r=10, t=40, b=10),
                           legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                           barmode=barmode)
        return fig

    grid = [
        line_fig("👥 Çalışan Sayısı", {'Çalışan Sayısı': series('employees')}, ['#3b82f6'], 'num', 0),
        bar_fig("📊 Raporlu Oran", {'Raporlu Oran': series('devamsizlik')}, ['#6366f1'], 'percent'),
        bar_fig("💰 Net Kök Ücret & İşveren Maliyeti",
                {'Net Kök Ücret': series('netKokUcret'), 'İşveren Maliyeti': series('isverenMaliyet')},
                ['#22c55e', '#fb923c'], 'tl'),
        bar_fig("⏱️ FM Saat", {'FM Saat': series('fmSaat')}, ['#fb923c'], 'num', 0),
        bar_fig("💸 FM (Net TL)", {'FM (Net TL)': series('fmTlMaliyet')}, ['#ef4444'], 'tl'),
        bar_fig("📅 İzin Gün Bakiyesi", {'İzin Gün': series('izinGun')}, ['#14b8a6'], 'num', 1),
        bar_fig("💎 İzin Ücreti (Net TL)", {'İzin Ücreti': series('izinUcret')}, ['#8b5cf6'], 'tl'),
        bar_fig("🏷️ Kıdem & İhbar Tazminatı (Net TL)",
                {'Kıdem Tazminatı': series('kidemTazminati'), 'İhbar Tazminatı': series('ihbarTazminati')},
                ['#64748b', '#ef4444'], 'tl'),
        line_fig("🧮 Kişi Başı Ortalama Maaş (Net TL)", {'Kişi Başı Ort. Maaş': series('kisiBasiOrt')}, ['#22c55e'], 'tl'),
        line_fig("👩 Kadın Oranı", {'Kadın Oranı': series('kadinOrani')}, ['#ec4899'], 'percent'),
        bar_fig("⏳ İlk 6 Ay İşten Ayrılma Oranı", {'İlk 6 Ay Ayrılma': series('ilk6ayOrani')}, ['#ef4444'], 'percent'),
        bar_fig("🔁 Ay İçi İşe Giren & İşten Ayrılan",
                {'Giriş': series('aySekiceGiris'), 'Çıkış': series('aySekiceCikis')}, ['#22c55e', '#ef4444'], 'num', 0),
    ]

    for i in range(0, len(grid), 2):
        gcol1, gcol2 = st.columns(2)
        gcol1.plotly_chart(grid[i], use_container_width=True)
        if i + 1 < len(grid):
            gcol2.plotly_chart(grid[i + 1], use_container_width=True)

    st.markdown("---")

    # ----- LOKASYON BAZINDA TABLOLAR VE GRAFİKLER (EN ÇOK MESAİ TABLOSUNDAN ÖNCE) -----
    st.subheader(f"📍 {selected_month} Ayı - Lokasyon Bazında Çalışan Sayısı ve Net Kök Ücret")

    if selected_month in lokasyon_cal.columns and selected_month in lokasyon_ucret.columns:
        # Grafik için veri (Toplam hariç)
        df_lok_chart = pd.DataFrame({
            'Lokasyon': lokasyon_cal.index,
            'Çalışan Sayısı': lokasyon_cal[selected_month],
            'Net Kök Ücret': lokasyon_ucret[selected_month]
        })

        # Tablo için veri (Toplam dahil)
        df_lok_table = df_lok_chart.copy()
        toplam_calisan = df_lok_table['Çalışan Sayısı'].sum()
        toplam_ucret = df_lok_table['Net Kök Ücret'].sum()
        df_lok_table.loc['Toplam'] = ['Toplam', toplam_calisan, toplam_ucret]

        # Tablo formatlama
        df_lok_table['Çalışan Sayısı'] = df_lok_table['Çalışan Sayısı'].apply(
            lambda x: format_number(x, 0) if isinstance(x, (int, float)) else x
        )
        df_lok_table['Net Kök Ücret'] = df_lok_table['Net Kök Ücret'].apply(
            lambda x: format_tl_no_decimal(x) if isinstance(x, (int, float)) else x
        )

        # Grafikler (2 sütun)
        col_chart1, col_chart2 = st.columns(2)

        # Çalışan Sayısı Grafiği
        fig_calisan = px.bar(
            df_lok_chart, x='Lokasyon', y='Çalışan Sayısı',
            title=f"{selected_month} - Lokasyon Bazında Çalışan Sayısı",
            text='Çalışan Sayısı', color='Lokasyon',
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig_calisan.update_traces(texttemplate='%{text:.0f}', textposition='outside')
        fig_calisan.update_layout(showlegend=False, height=400, xaxis_title='', yaxis_title='Çalışan Sayısı')
        col_chart1.plotly_chart(fig_calisan, use_container_width=True)

        # Net Kök Ücret Grafiği
        fig_ucret = px.bar(
            df_lok_chart, x='Lokasyon', y='Net Kök Ücret',
            title=f"{selected_month} - Lokasyon Bazında Net Kök Ücret",
            text='Net Kök Ücret', color='Lokasyon',
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        fig_ucret.update_traces(texttemplate='%{text:,.0f} TL', textposition='outside')
        fig_ucret.update_layout(showlegend=False, height=400, xaxis_title='', yaxis_title='Net Kök Ücret (TL)')
        col_chart2.plotly_chart(fig_ucret, use_container_width=True)

        # Tablo
        st.dataframe(df_lok_table, use_container_width=True, hide_index=True)
    else:
        st.info("Seçilen ay için lokasyon verisi bulunamadı.")

    st.markdown("---")

    # ----- EN ÇOK MESAİ YAPAN 10 KİŞİ -----
    st.subheader(f"🏆 {selected_month} Ayında En Çok Mesai Yapan 10 Kişi (Haziran Servis)")
    hz_fm_df = fm_yapan_df[fm_yapan_df['Şirket'] == COMPANY].copy() if 'Şirket' in fm_yapan_df.columns else fm_yapan_df.copy()
    if selected_month in hz_fm_df.columns and not hz_fm_df.empty:
        cols_needed = [c for c in ['Adı Soyadı', 'Lokasyon', 'Organizasyon', 'Departman'] if c in hz_fm_df.columns]
        top10 = hz_fm_df[cols_needed + [selected_month]].copy()
        top10 = top10.sort_values(by=selected_month, ascending=False).head(10)
        top10 = top10.rename(columns={selected_month: 'FM Saat'})
        top10['FM Saat'] = top10['FM Saat'].apply(lambda x: format_number(x, 1))
        top10.insert(0, 'Sıra', range(1, len(top10) + 1))
        st.dataframe(top10, use_container_width=True, hide_index=True)
    else:
        st.info("Bu ay için mesai verisi bulunamadı.")

    st.markdown("---")

    # ----- DETAY TABLOSU -----
    st.subheader(f"📋 {selected_month} - Haziran Servis Detayları")
    detail = {
        'Çalışan': format_number(d('employees'), 0),
        'Raporlu Oran %': format_percent(d('devamsizlik')),
        'Küm. Turnover %': format_percent(d('turnoverKumulatif')),
        'Küm. Gön. %': format_percent(d('turnoverGonulluKumulatif')),
        'Aylık Turnover %': format_percent(d('turnoverAylik')),
        'Aylık Gön. %': format_percent(d('turnoverGonulluAylik')),
        'Net Kök Ücret': format_tl(d('netKokUcret')),
        'İşveren Maliyeti': format_tl(d('isverenMaliyet')),
        'FM Saat': format_number(d('fmSaat'), 1),
        'FM (Net TL)': format_tl(d('fmTlMaliyet')),
        'İzin Gün Bakiyesi': format_number(d('izinGun'), 1),
        'İzin Ücreti (Net TL)': format_tl(d('izinUcret')),
        'Kişi Başı Ort. Maaş (Net TL)': format_tl(d('kisiBasiOrt')),
        'Yıllık Ort. Maaş (Net TL)': format_tl(d('yillikOrtMaas')),
        'Kıdem Tazminatı (Net TL)': format_tl(d('kidemTazminati')),
        'İhbar Tazminatı (Net TL)': format_tl(d('ihbarTazminati')),
        'Kadın Oranı %': format_percent(d('kadinOrani')),
        'İlk 6 Ay Ayrılma %': format_percent(d('ilk6ayOrani')),
        'Ay İçi İşe Giren': format_number(d('aySekiceGiris'), 0),
        'Ay İçi İşten Ayrılan': format_number(d('aySekiceCikis'), 0),
    }
    detail_df = pd.DataFrame([detail])
    st.dataframe(detail_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()