import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="Haziran Servis Dashboard")

# ----- DOSYA YOLU VE ŞİFRE AYARLARI -----
DATA_DIR = "data"
DATA_PATH = os.path.join(DATA_DIR, "son_veri.xlsx")
os.makedirs(DATA_DIR, exist_ok=True)  # Klasör yoksa otomatik oluştur

# Şifre (basit tutuyoruz, isterseniz değiştirin)
UPLOAD_PASSWORD = "ik2026"

# ----- FORMAT YARDIMCILARI (eski kodla aynı) -----
def format_tl(value):
    if value is None:
        return "0TL"
    try:
        formatted = f"{value:,.2f}"
        formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{formatted}TL"
    except:
        return f"{value:.2f}TL".replace(".", ",")

def format_tl_no_decimal(value):
    if value is None:
        return "0TL"
    try:
        formatted = f"{value:,.0f}"
        formatted = formatted.replace(",", ".")
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
        formatted = f"{value:,.{decimals}f}"
        formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
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
        'HAZIRAN': 'Haziran Servis',
        'Haziran': 'Haziran Servis'
    }
    return replacements.get(name, name)

def clean_columns(df):
    df.columns = [str(col).strip() for col in df.columns]
    return df

MONTHS = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz']

def get_month_cols(df):
    month_cols = []
    for col in df.columns:
        col_clean = str(col).strip()
        for m in MONTHS:
            if col_clean.lower() == m.lower():
                month_cols.append(col)
                break
    return month_cols

# ----- VERİ YÜKLEME (SADECE HAZİRAN SERVİS) -----
@st.cache_data
def load_haziran_data():
    if not os.path.exists(DATA_PATH):
        st.error("Veri dosyası bulunamadı! Lütfen soldaki 'Veriyi Güncelle' panelinden Excel dosyasını yükleyin.")
        return None

    # 1. Genel Turnover
    df_gt = pd.read_excel(DATA_PATH, sheet_name='genel.turnover', header=0)
    df_gt = clean_columns(df_gt)
    df_gt['Şirket'] = df_gt.iloc[:, 0].apply(normalize_company_name)
    df_gt = df_gt.set_index('Şirket')
    month_cols = get_month_cols(df_gt)
    haziran_gt = df_gt.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_gt.index else pd.Series(0, index=month_cols)

    # 2. Gönüllü Turnover
    df_gon = pd.read_excel(DATA_PATH, sheet_name='gonullu.turnover', header=0)
    df_gon = clean_columns(df_gon)
    df_gon['Şirket'] = df_gon.iloc[:, 0].apply(normalize_company_name)
    df_gon = df_gon.set_index('Şirket')
    month_cols = get_month_cols(df_gon)
    haziran_gon = df_gon.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_gon.index else pd.Series(0, index=month_cols)

    # 3. Rapor Oranı
    df_rapor = pd.read_excel(DATA_PATH, sheet_name='rapor_oran', header=0)
    df_rapor = clean_columns(df_rapor)
    df_rapor['Şirket'] = df_rapor.iloc[:, 0].apply(normalize_company_name)
    df_rapor = df_rapor.set_index('Şirket')
    month_cols = get_month_cols(df_rapor)
    haziran_rapor = df_rapor.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_rapor.index else pd.Series(0, index=month_cols)

    # 4. Çalışan Sayısı
    df_calisan = pd.read_excel(DATA_PATH, sheet_name='calisan.sayisi', header=0)
    df_calisan = clean_columns(df_calisan)
    df_calisan['Şirket'] = df_calisan.iloc[:, 0].apply(normalize_company_name)
    df_calisan = df_calisan.set_index('Şirket')
    month_cols = get_month_cols(df_calisan)
    haziran_calisan = df_calisan.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_calisan.index else pd.Series(0, index=month_cols)

    # 5. Net Kök Ücret
    df_net = pd.read_excel(DATA_PATH, sheet_name='maliyet', header=0)
    df_net = clean_columns(df_net)
    df_net['Şirket'] = df_net.iloc[:, 0].apply(normalize_company_name)
    df_net = df_net.set_index('Şirket')
    month_cols = get_month_cols(df_net)
    haziran_net = df_net.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_net.index else pd.Series(0, index=month_cols)

    # 6. İşveren Maliyeti
    df_isv = pd.read_excel(DATA_PATH, sheet_name='isveren.maliyet', header=0)
    df_isv = clean_columns(df_isv)
    df_isv['Şirket'] = df_isv.iloc[:, 0].apply(normalize_company_name)
    df_isv = df_isv.set_index('Şirket')
    month_cols = get_month_cols(df_isv)
    haziran_isv = df_isv.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_isv.index else pd.Series(0, index=month_cols)

    # 7. FM Saat
    df_fm_saat = pd.read_excel(DATA_PATH, sheet_name='fm.saat', header=0)
    df_fm_saat = clean_columns(df_fm_saat)
    df_fm_saat['Şirket'] = df_fm_saat.iloc[:, 0].apply(normalize_company_name)
    df_fm_saat = df_fm_saat.set_index('Şirket')
    month_cols = get_month_cols(df_fm_saat)
    haziran_fm_saat = df_fm_saat.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_fm_saat.index else pd.Series(0, index=month_cols)

    # 8. FM TL Maliyet
    df_fm_tl = pd.read_excel(DATA_PATH, sheet_name='fm.maliyet', header=0)
    df_fm_tl = clean_columns(df_fm_tl)
    df_fm_tl['Şirket'] = df_fm_tl.iloc[:, 0].apply(normalize_company_name)
    df_fm_tl = df_fm_tl.set_index('Şirket')
    month_cols = get_month_cols(df_fm_tl)
    haziran_fm_tl = df_fm_tl.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_fm_tl.index else pd.Series(0, index=month_cols)

    # 9. İzin Gün
    df_izin_gun = pd.read_excel(DATA_PATH, sheet_name='izin_gun', header=0)
    df_izin_gun = clean_columns(df_izin_gun)
    df_izin_gun['Şirket'] = df_izin_gun.iloc[:, 0].apply(normalize_company_name)
    df_izin_gun = df_izin_gun.set_index('Şirket')
    month_cols = get_month_cols(df_izin_gun)
    haziran_izin_gun = df_izin_gun.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_izin_gun.index else pd.Series(0, index=month_cols)

    # 10. İzin Ücreti
    df_izin_ucret = pd.read_excel(DATA_PATH, sheet_name='izin_ucret', header=0)
    df_izin_ucret = clean_columns(df_izin_ucret)
    df_izin_ucret['Şirket'] = df_izin_ucret.iloc[:, 0].apply(normalize_company_name)
    df_izin_ucret = df_izin_ucret.set_index('Şirket')
    month_cols = get_month_cols(df_izin_ucret)
    haziran_izin_ucret = df_izin_ucret.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_izin_ucret.index else pd.Series(0, index=month_cols)

    # 11. Kıdem Tazminatı
    df_kidem = pd.read_excel(DATA_PATH, sheet_name='kidem.tazminati', header=0)
    df_kidem = clean_columns(df_kidem)
    df_kidem['Şirket'] = df_kidem.iloc[:, 0].apply(normalize_company_name)
    df_kidem = df_kidem.set_index('Şirket')
    month_cols = get_month_cols(df_kidem)
    is_total = df_kidem.index.astype(str).str.strip().str.upper() == 'TOPLAM'
    df_kidem = df_kidem[~is_total]
    haziran_kidem = df_kidem.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_kidem.index else pd.Series(0, index=month_cols)

    # 12. İhbar Tazminatı
    df_ihbar = pd.read_excel(DATA_PATH, sheet_name='ihbar.tazminati', header=0)
    df_ihbar = clean_columns(df_ihbar)
    df_ihbar['Şirket'] = df_ihbar.iloc[:, 0].apply(normalize_company_name)
    df_ihbar = df_ihbar.set_index('Şirket')
    month_cols = get_month_cols(df_ihbar)
    is_total = df_ihbar.index.astype(str).str.strip().str.upper() == 'TOPLAM'
    df_ihbar = df_ihbar[~is_total]
    haziran_ihbar = df_ihbar.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_ihbar.index else pd.Series(0, index=month_cols)

    # 13. Kişi Başı Ortalama Maaş
    df_kisi = pd.read_excel(DATA_PATH, sheet_name='kisi.basi.ort', header=0)
    df_kisi = clean_columns(df_kisi)
    df_kisi['Şirket'] = df_kisi.iloc[:, 0].apply(normalize_company_name)
    df_kisi = df_kisi.set_index('Şirket')
    month_cols = get_month_cols(df_kisi)
    is_genel = df_kisi.index.astype(str).str.strip().str.upper().str.contains('GENEL')
    df_kisi = df_kisi[~is_genel]
    haziran_kisi = df_kisi.loc['Haziran Servis', month_cols] if 'Haziran Servis' in df_kisi.index else pd.Series(0, index=month_cols)

    # 14. Aylık FM Yapan (Haziran Servis filtresi)
    df_fm_yapan = pd.read_excel(DATA_PATH, sheet_name='aylik.fm.yapan', header=0)
    df_fm_yapan = clean_columns(df_fm_yapan)
    if 'Şirket' in df_fm_yapan.columns:
        df_fm_yapan['Şirket'] = df_fm_yapan['Şirket'].apply(normalize_company_name)
    df_fm_yapan = df_fm_yapan[df_fm_yapan['Şirket'] == 'Haziran Servis']
    for m in get_month_cols(df_fm_yapan):
        df_fm_yapan[m] = pd.to_numeric(df_fm_yapan[m], errors='coerce').fillna(0)

    return {
        'gt': haziran_gt,
        'gon': haziran_gon,
        'rapor': haziran_rapor,
        'calisan': haziran_calisan,
        'net': haziran_net,
        'isveren': haziran_isv,
        'fm_saat': haziran_fm_saat,
        'fm_tl': haziran_fm_tl,
        'izin_gun': haziran_izin_gun,
        'izin_ucret': haziran_izin_ucret,
        'kidem': haziran_kidem,
        'ihbar': haziran_ihbar,
        'kisi_basi': haziran_kisi,
        'fm_yapan': df_fm_yapan
    }

# ----- ANA UYGULAMA -----
def main():
    st.title("📊 Haziran Servis Özel Dashboard")

    # ----- VERİ GÜNCELLEME PANELİ (SIDEBAR) -----
    with st.sidebar:
        with st.expander("🔒 Veriyi Güncelle", expanded=not os.path.exists(DATA_PATH)):
            pwd = st.text_input("Şifre", type="password")
            new_file = st.file_uploader("Yeni Excel dosyası", type=["xlsx"], key="haziran_uploader")
            if new_file is not None:
                if pwd == UPLOAD_PASSWORD:
                    with open(DATA_PATH, "wb") as f:
                        f.write(new_file.getbuffer())
                    st.cache_data.clear()
                    st.success("✅ Veri başarıyla güncellendi!")
                    st.rerun()
                else:
                    st.error("❌ Şifre yanlış!")

    st.markdown("---")

    # Veri kontrolü
    if not os.path.exists(DATA_PATH):
        st.info("📂 Henüz veri yüklenmedi. Soldaki '🔒 Veriyi Güncelle' panelinden Excel dosyasını yükleyin.")
        st.stop()

    # Veriyi yükle
    data = load_haziran_data()
    if data is None:
        st.stop()

    # Ay seçimi
    selected_month = st.selectbox("📅 Ay Seçin", MONTHS, index=MONTHS.index("Temmuz"))
    month_idx = MONTHS.index(selected_month)
    prev_month = MONTHS[month_idx - 1] if month_idx > 0 else None

    def get_val(series, month):
        return series.get(month, 0) if month in series.index else 0

    cur_gt = get_val(data['gt'], selected_month)
    prev_gt = get_val(data['gt'], prev_month) if prev_month else None

    cur_gon = get_val(data['gon'], selected_month)
    cur_rapor = get_val(data['rapor'], selected_month) * 100
    prev_rapor = get_val(data['rapor'], prev_month) * 100 if prev_month else None

    cur_calisan = get_val(data['calisan'], selected_month)
    prev_calisan = get_val(data['calisan'], prev_month) if prev_month else None

    cur_net = get_val(data['net'], selected_month)
    prev_net = get_val(data['net'], prev_month) if prev_month else None

    cur_isveren = get_val(data['isveren'], selected_month)
    prev_isveren = get_val(data['isveren'], prev_month) if prev_month else None

    cur_fm_saat = get_val(data['fm_saat'], selected_month)
    prev_fm_saat = get_val(data['fm_saat'], prev_month) if prev_month else None

    cur_fm_tl = get_val(data['fm_tl'], selected_month)
    prev_fm_tl = get_val(data['fm_tl'], prev_month) if prev_month else None

    cur_izin_gun = get_val(data['izin_gun'], selected_month)
    prev_izin_gun = get_val(data['izin_gun'], prev_month) if prev_month else None

    cur_izin_ucret = get_val(data['izin_ucret'], selected_month)
    prev_izin_ucret = get_val(data['izin_ucret'], prev_month) if prev_month else None

    cur_kidem = get_val(data['kidem'], selected_month)
    prev_kidem = get_val(data['kidem'], prev_month) if prev_month else None

    cur_ihbar = get_val(data['ihbar'], selected_month)
    prev_ihbar = get_val(data['ihbar'], prev_month) if prev_month else None

    cur_kisi = get_val(data['kisi_basi'], selected_month)
    prev_kisi = get_val(data['kisi_basi'], prev_month) if prev_month else None

    # KPI Kartları
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("👥 Çalışan", format_number(cur_calisan, 0),
                delta=format_delta_number(calc_diff(cur_calisan, prev_calisan), 0))
    col2.metric("📊 Raporlu Oran", format_percent(cur_rapor),
                delta=format_delta_percent(calc_diff(cur_rapor, prev_rapor)))
    col3.metric("💼 İşveren Maliyeti", format_tl(cur_isveren),
                delta=format_delta_tl(calc_diff(cur_isveren, prev_isveren)))
    col4.metric("💰 Net Kök Ücret", format_tl(cur_net),
                delta=format_delta_tl(calc_diff(cur_net, prev_net)))
    col5.metric("⏱️ FM Saat", format_number(cur_fm_saat, 1),
                delta=format_delta_number(calc_diff(cur_fm_saat, prev_fm_saat), 1))
    col6.metric("💸 FM (Net TL)", format_tl(cur_fm_tl),
                delta=format_delta_tl(calc_diff(cur_fm_tl, prev_fm_tl)))

    col7, col8, col9, col10, col11 = st.columns(5)
    col7.metric("📅 İzin Gün", format_number(cur_izin_gun, 1),
                delta=format_delta_number(calc_diff(cur_izin_gun, prev_izin_gun), 1))
    col8.metric("💎 İzin Ücreti", format_tl(cur_izin_ucret),
                delta=format_delta_tl(calc_diff(cur_izin_ucret, prev_izin_ucret)))
    col9.metric("🏷️ Kıdem Tazminatı", format_tl(cur_kidem),
                delta=format_delta_tl(calc_diff(cur_kidem, prev_kidem)))
    col10.metric("📨 İhbar Tazminatı", format_tl(cur_ihbar),
                 delta=format_delta_tl(calc_diff(cur_ihbar, prev_ihbar)))
    col11.metric("🧮 Kişi Başı Ort. Maaş", format_tl(cur_kisi),
                 delta=format_delta_tl(calc_diff(cur_kisi, prev_kisi)))

    st.markdown("---")

    # ----- EN ÇOK MESAİ YAPAN 10 KİŞİ -----
    st.subheader(f"🏆 {selected_month} Ayında En Çok Mesai Yapan 10 Kişi (Haziran Servis)")

    fm_df = data['fm_yapan']
    if selected_month in fm_df.columns and not fm_df.empty:
        cols = ['Adı Soyadı', 'Lokasyon', 'Organizasyon', 'Departman', selected_month]
        available = [c for c in cols if c in fm_df.columns]
        if len(available) < 5:
            st.warning("Bazı sütunlar eksik, mevcut sütunlar gösteriliyor.")
            available = [c for c in fm_df.columns if c in ['Adı Soyadı', 'Lokasyon', 'Organizasyon', 'Departman'] or c == selected_month]
        top10 = fm_df[available].copy()
        top10 = top10.sort_values(by=selected_month, ascending=False).head(10)
        top10 = top10.rename(columns={selected_month: 'FM Saat'})
        top10['FM Saat'] = top10['FM Saat'].apply(lambda x: format_number(x, 1))
        top10.insert(0, 'Sıra', range(1, len(top10)+1))
        st.dataframe(top10, use_container_width=True, hide_index=True)
    else:
        st.info("Bu ay için mesai verisi bulunamadı.")

    st.markdown("---")

    # ----- DETAY TABLOSU -----
    st.subheader(f"📋 {selected_month} - Haziran Servis Detayları")
    detail = {
        'Şirket': 'Haziran Servis',
        'Çalışan': format_number(cur_calisan, 0),
        'Raporlu Oran %': format_percent(cur_rapor),
        'Turnover Toplam %': format_percent(cur_gt * 100),
        'Turnover Gönüllü %': format_percent(cur_gon * 100),
        'Net Kök Ücret': format_tl(cur_net),
        'İşveren Maliyeti': format_tl(cur_isveren),
        'FM Saat': format_number(cur_fm_saat, 1),
        'FM (Net TL)': format_tl(cur_fm_tl),
        'İzin Gün Bakiyesi': format_number(cur_izin_gun, 1),
        'İzin Ücreti (Net TL)': format_tl(cur_izin_ucret),
        'Kişi Başı Ort. Maaş (Net TL)': format_tl(cur_kisi),
        'Kıdem Tazminatı': format_tl(cur_kidem),
        'İhbar Tazminatı': format_tl(cur_ihbar),
    }
    detail_df = pd.DataFrame([detail])
    st.dataframe(detail_df, use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()