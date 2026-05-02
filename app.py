import streamlit as st
import pandas as pd
import xlsxwriter
import re
import datetime
import io

# --- GOOGLE SHEETS BAGLANTISI ---
SHEET_ID = "1Polxg5n-J0VueJifvjlgoGvXITVvnBbJpgSjHJ6imYY"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"

st.set_page_config(page_title="IHMC 2026 Matbaa", page_icon="🖨️", layout="wide")

# --- YARDIMCI FONKSIYONLAR ---
def clean_salon(sal):
    if isinstance(sal, pd.Timestamp) or hasattr(sal, 'month'):
        if sal.month == 8 and sal.day == 30: return '30 Ağustos Salonu'
            
    n = str(sal).upper().strip()
    if '30 A' in n or '08-30' in n: return '30 Ağustos Salonu'
    if 'MALAZGIRT' in n or 'MALAZGİRT' in n: return 'Malazgirt Salonu'
    if 'ELIF' in n or 'ELİF' in n: return 'Doç. Dr. Elif Kaya Salonu'
    if 'AHMET' in n or 'EKIZER' in n or 'EKİZER' in n: return 'Dr. Ahmet Ekizer Salonu'
    
    # TEAMS ODALARI (Excel'de yanlışlıkla Zoom yazsa bile Teams'e çevirir)
    if 'TEAMS 1' in n or 'ZOOM 1' in n or 'ODA 1' in n: return 'Teams Oda 1'
    if 'TEAMS 2' in n or 'ZOOM 2' in n or 'ODA 2' in n: return 'Teams Oda 2'
    if 'TEAMS 3' in n or 'ZOOM 3' in n or 'ODA 3' in n: return 'Teams Oda 3'
    
    return str(sal).strip()

def parse_zaman(metin):
    match = re.search(r'(\d{2}\.\d{2}\.\d{4}).*?(\d{2}[:.]\d{2})', str(metin))
    if match:
        date_str = match.group(1)
        time_str = match.group(2).replace('.', ':')
        try:
            return datetime.datetime.strptime(f"{date_str} {time_str}", "%d.%m.%Y %H:%M")
        except: pass
    return datetime.datetime(2099, 1, 1)

# --- ANA EXCEL OLUSTURMA MOTORU ---
def excel_bas(df_bildiriler, df_ozel, df_mod, tip):
    prog = {}
    tum_zamanlar = set()
    tum_salonlar = set()
    
    for index, r in df_bildiriler.iterrows():
        b_adi = str(r['Bildiri_Adi'])
        sunucu = str(r['sunucu'])
        yazarlar = str(r['yazarlar'])
        konu = str(r['konu'])
        
        kz = str(r['Gun_ve_Saat']).replace("Persembe", "Perşembe") 
        ks = clean_salon(r['Salon'])
        sid = str(r['Oturum_ID'])
        
        tum_zamanlar.add(kz)
        tum_salonlar.add(ks)
        
        anahtar = (kz, ks)
        if anahtar not in prog: prog[anahtar] = []
        prog[anahtar].append({'b': b_adi, 'y': yazarlar, 's': sunucu, 'k': konu, 'sid': sid})

    soft_palette = ['#DDEBF7', '#FCE4D6', '#E2EFDA', '#FFF2CC', '#E6E6FA', '#F2F2F2']
    color_index = 0

    gosterilecek_ozel_etkinlikler = []
    if not df_ozel.empty:
        for (sira, sal), grup in df_ozel.groupby(['oturum_sirasi', 'salon'], sort=False):
            ks = clean_salon(sal)
            tum_salonlar.add(ks)
            
            renk_temasi = soft_palette[color_index % len(soft_palette)]
            color_index += 1
            
            ts = str(grup.iloc[0].get('tarih_saat', '-')).strip().replace("Persembe", "Perşembe")
            t_event = "07.05.2026 Perşembe" if "07" in ts else ("08.05.2026 Cuma" if "08" in ts else "Bilinmeyen Gun")
            m = re.search(r'(\d{2}[.:]\d{2})\s*-\s*(\d{2}[.:]\d{2})', ts)
            sa_event = f"{m.group(1).replace('.',':')}-{m.group(2).replace('.',':')}" if m else ts
            zaman_key = f"{t_event} | {sa_event}"
            
            tum_zamanlar.add(zaman_key)
            gosterilecek_ozel_etkinlikler.append((sira, ks, grup, renk_temasi, zaman_key))

    musait_oturumlar_mod = list(prog.keys())
    musait_oturumlar_deg = list(prog.keys())
    mod_atamalari = {k: {'mod': 'Daha Sonra İlan Edilecektir', 'deg': 'Daha Sonra İlan Edilecektir'} for k in prog.keys()}
    serbest_mods, serbest_degs = [], []
    
    if not df_mod.empty:
        for _, r in df_mod.iterrows():
            is_online_mod = str(r.get('Online', '-')).strip().lower() in ['evet', 'e', 'yes', '1', 'var', 'online', 'true']
            if tip == "ONLINE" and not is_online_mod: continue 
            if tip == "YUZYUZE" and is_online_mod: continue
            
            m_gun = str(r.get('zorunlu_gun', '-')).strip()
            m_oturum = str(r.get('zorunlu_oturum', '-')).strip()
            m_saat = str(r.get('zorunlu_saat', '-')).strip()
            gorev_sutunu = str(r.get('Gorev', '-')).strip().lower()
            
            mod_metni = f"{r.get('unvan_ad_soyad', '-')} ({r.get('kurum', '')})" if str(r.get('kurum', '-')) not in ['-', 'nan', ''] else str(r.get('unvan_ad_soyad', '-'))
            
            is_deg = 'deg' in gorev_sutunu
            role_key = 'deg' if is_deg else 'mod'
            hedef_musait_liste = musait_oturumlar_deg if is_deg else musait_oturumlar_mod
            hedef_serbest_liste = serbest_degs if is_deg else serbest_mods

            atandi_mi = False
            if m_gun != '-' or m_oturum != '-' or m_saat != '-':
                for sess in hedef_musait_liste:
                    zaman_metni = sess[0]
                    oturum_id_metni = prog[sess][0]['sid'] if prog[sess] else "-"
                    if (m_gun == '-' or m_gun in zaman_metni) and \
                       (m_oturum == '-' or m_oturum in zaman_metni or m_oturum == oturum_id_metni) and \
                       (m_saat == '-' or m_saat in zaman_metni):
                        mod_atamalari[sess][role_key] = mod_metni
                        hedef_musait_liste.remove(sess)
                        atandi_mi = True
                        break
            if not atandi_mi: hedef_serbest_liste.append(mod_metni)
                
        for mod_metni in serbest_mods:
            if musait_oturumlar_mod: mod_atamalari[musait_oturumlar_mod.pop(0)]['mod'] = mod_metni
        for deg_metni in serbest_degs:
            if musait_oturumlar_deg: mod_atamalari[musait_oturumlar_deg.pop(0)]['deg'] = deg_metni

    gun_istatistikleri = {}
    for (oturum_zaman, sal), bilds in prog.items():
        t = oturum_zaman.split(' | ')[0] if ' | ' in oturum_zaman else oturum_zaman
        if t not in gun_istatistikleri: gun_istatistikleri[t] = {'oturum_sayisi': 0, 'bildiri_sayisi': 0}
        gun_istatistikleri[t]['oturum_sayisi'] += 1
        gun_istatistikleri[t]['bildiri_sayisi'] += len(bilds)

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    map_data = {}
    for (kz, ks), bilds in prog.items():
        if (kz, ks) not in map_data: map_data[(kz, ks)] = {'session_count': 0, 'event_name': None, 'event_color': None}
        map_data[(kz, ks)]['session_count'] = len(bilds)

    for sira, ks, grup, renk, zaman_key in gosterilecek_ozel_etkinlikler:
        ana_baslik = str(grup.iloc[0].get('ana_baslik', 'Etkinlik')).strip()
        if (zaman_key, ks) not in map_data: map_data[(zaman_key, ks)] = {'session_count': 0, 'event_name': None, 'event_color': None}
        map_data[(zaman_key, ks)]['event_name'] = ana_baslik
        map_data[(zaman_key, ks)]['event_color'] = renk

    list_zamanlar = sorted(list(tum_zamanlar), key=parse_zaman)
    list_salonlar = sorted(list(tum_salonlar))

    ws_harita = workbook.add_worksheet('Oturum_Haritasi')
    f_map_baslik = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    f_map_saat = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'left'})
    f_map_bos = workbook.add_format({'bg_color': '#F2F2F2', 'font_color': '#A6A6A6', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    f_map_dolu = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'font_color': '#375623', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    f_map_cakisma = workbook.add_format({'bold': True, 'bg_color': '#FF0000', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    
    dinamik_renk_cache = {}

    ws_harita.set_column(0, 0, 35)
    ws_harita.write(0, 0, "SAAT / OTURUM", f_map_baslik)
    for col_idx, sal in enumerate(list_salonlar, 1):
        ws_harita.set_column(col_idx, col_idx, 25)
        ws_harita.write(0, col_idx, sal, f_map_baslik)

    for row_idx, z_key in enumerate(list_zamanlar, 1):
        ws_harita.write(row_idx, 0, z_key, f_map_saat)
        for col_idx, sal in enumerate(list_salonlar, 1):
            hucre = map_data.get((z_key, sal), {'session_count': 0, 'event_name': None})
            s_count = hucre['session_count']
            e_name = hucre['event_name']
            e_color = hucre.get('event_color', '#FFD966')
            
            if s_count > 0 and e_name:
                metin = f"ÇAKIŞMA!\n[{s_count} Bildiri] VE [{e_name}]"
                ws_harita.write(row_idx, col_idx, metin, f_map_cakisma)
                ws_harita.set_row(row_idx, 45)
            elif s_count > 0:
                ws_harita.write(row_idx, col_idx, f"{s_count} Bildiri Var", f_map_dolu)
            elif e_name:
                if e_color not in dinamik_renk_cache:
                    dinamik_renk_cache[e_color] = workbook.add_format({'bold': True, 'bg_color': e_color, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                ws_harita.write(row_idx, col_idx, f"[ETKİNLİK]\n{e_name}", dinamik_renk_cache[e_color])
                ws_harita.set_row(row_idx, 40)
            else:
                ws_harita.write(row_idx, col_idx, "BOŞ", f_map_bos)

    worksheet = workbook.add_worksheet('Kesin_Program')
    worksheet.set_column('A:A', 50); worksheet.set_column('B:B', 60) 
    
    fmt = {
        'ayirici': workbook.add_format({'bg_color': '#000000', 'font_color': '#FFFFFF', 'bold': True, 'align': 'center', 'valign': 'vcenter'}),
        'mavi_gun': workbook.add_format({'bold': True, 'bg_color': '#B4C6E7', 'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True}),
        'odul_sari': workbook.add_format({'bg_color': '#FFF2CC', 'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True}),
        'istatistik_yesil': workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'font_color': '#375623', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True}),
        'saat_baslik': workbook.add_format({'bold': True, 'bg_color': '#FFD966', 'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True}),
        'salon_baslik': workbook.add_format({'bold': True, 'bg_color': '#FFE699', 'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True}),
        'bildiri_tek': workbook.add_format({'bg_color': '#FFFFFF', 'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True}),
        'bildiri_tek_u': workbook.add_format({'bold': True, 'underline': True, 'bg_color': '#FFFFFF', 'border': 1, 'text_wrap': True}),
        'bildiri_tek_u_only': workbook.add_format({'bold': True, 'underline': True}),
        'bildiri_cift': workbook.add_format({'bg_color': '#F2F2F2', 'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True}),
        'bildiri_cift_u': workbook.add_format({'bold': True, 'underline': True, 'bg_color': '#F2F2F2', 'border': 1, 'text_wrap': True}),
        'bildiri_cift_u_only': workbook.add_format({'bold': True, 'underline': True})
    }

    satir_no = 0
    if len(gosterilecek_ozel_etkinlikler) > 0:
        worksheet.merge_range(satir_no, 0, satir_no, 1, "--- ÖZEL ETKİNLİKLER ---", fmt['ayirici']); satir_no += 1
        for sira, ks, grup, renk_temasi, z_key in gosterilecek_ozel_etkinlikler:
            ts = str(grup.iloc[0].get('tarih_saat', '-')).strip().replace("07.05.2026", "07.05.2026 Perşembe /").replace("08.05.2026", "08.05.2026 Cuma /")
            if renk_temasi not in dinamik_renk_cache:
                dinamik_renk_cache[renk_temasi] = workbook.add_format({'bold': True, 'bg_color': renk_temasi, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
            dinamik_f = dinamik_renk_cache[renk_temasi]
            worksheet.merge_range(satir_no, 0, satir_no, 1, f"{ts} | HALL: {ks}", dinamik_f); satir_no += 1
            for index, r in grup.iterrows():
                metin = "\n".join([str(r.get(c, '-')).strip() for c in ['ana_baslik', 'alt_baslik', 'sol_metin', 'sag_metin'] if str(r.get(c, '-')).strip() not in ['-', 'nan', '']])
                worksheet.merge_range(satir_no, 0, satir_no, 1, metin, dinamik_f); worksheet.set_row(satir_no, 50); satir_no += 1
            satir_no += 1

    worksheet.merge_range(satir_no, 0, satir_no, 1, "--- BİLİMSEL BİLDİRİ PROGRAMI ---", fmt['ayirici']); satir_no += 2

    mevcut_islenen_gun = ""
    sirali_oturumlar = sorted(prog.items(), key=lambda x: parse_zaman(x[0][0]))
    
    for (oturum_zaman, sal), bilds in sirali_oturumlar:
        parcalar = oturum_zaman.split(' | ')
        t = parcalar[0] if len(parcalar) > 0 else oturum_zaman
        sa = parcalar[1] if len(parcalar) > 1 else "-"
        sn = bilds[0]['sid'] if bilds else "-"
        
        if t != mevcut_islenen_gun:
            mevcut_islenen_gun = t
            worksheet.merge_range(satir_no, 0, satir_no, 1, f">>> {t.upper()} BİLİMSEL PROGRAMI <<<", fmt['mavi_gun'])
            worksheet.set_row(satir_no, 40); satir_no += 1
            
            odul_notu = "Her oturumdaki moderatör oturumu yönetecek ve Oturum Değerlendirici ile birbirinden bağımsız olarak EN İYİ BİLDİRİ (BEST PAPER) ÖDÜLLERi için bildiri sunumlarını değerlendireceklerdir."
            worksheet.merge_range(satir_no, 0, satir_no, 1, odul_notu, fmt['odul_sari'])
            worksheet.set_row(satir_no, 45); satir_no += 1 
            
            o_sayi = gun_istatistikleri.get(t, {}).get('oturum_sayisi', 0)
            b_sayi = gun_istatistikleri.get(t, {}).get('bildiri_sayisi', 0)
            istatistik_notu = f"Bugün toplam {o_sayi} adet bildiri sunum oturumu gerçekleşecek ve {b_sayi} adet bildiri sunulacaktır."
            worksheet.merge_range(satir_no, 0, satir_no, 1, istatistik_notu, fmt['istatistik_yesil'])
            worksheet.set_row(satir_no, 25); satir_no += 2
        
        mod_isim = mod_atamalari[(oturum_zaman, sal)]['mod']
        deg_isim = mod_atamalari[(oturum_zaman, sal)]['deg']
        oturum_metni = f"HALL: {sal}\n{sn}\n\n{bilds[0]['k'].upper()}\nModerator: {mod_isim}\nOturum Değerlendirici: {deg_isim}"
        
        worksheet.merge_range(satir_no, 0, satir_no, 1, f"{t} | {sa}", fmt['saat_baslik']); satir_no += 1
        worksheet.merge_range(satir_no, 0, satir_no, 1, oturum_metni, fmt['salon_baslik'])
        worksheet.set_row(satir_no, 100); satir_no += 1 
        
        for i, b in enumerate(bilds):
            yazarlar_metni, sunucu = b['y'], b['s']
            is_cift = (i % 2 == 1) 
            f_norm = fmt['bildiri_cift'] if is_cift else fmt['bildiri_tek']
            f_ul_cell = fmt['bildiri_cift_u'] if is_cift else fmt['bildiri_tek_u']
            f_ul_txt = fmt['bildiri_cift_u_only'] if is_cift else fmt['bildiri_tek_u_only']
            
            if sunucu == yazarlar_metni and sunucu not in ['-', 'nan', '']: 
                worksheet.write(satir_no, 0, yazarlar_metni, f_ul_cell)
            elif sunucu in yazarlar_metni and sunucu not in ['-', 'nan', '']:
                parcalar_yazar = yazarlar_metni.split(sunucu)
                rich_text = []
                if parcalar_yazar[0]: rich_text.extend([f_norm, parcalar_yazar[0]])
                rich_text.extend([f_ul_txt, sunucu])
                if len(parcalar_yazar) > 1 and parcalar_yazar[1]: rich_text.extend([f_norm, parcalar_yazar[1]])
                if len(rich_text) >= 2: worksheet.write_rich_string(satir_no, 0, *rich_text, f_norm)
                else: worksheet.write(satir_no, 0, yazarlar_metni, f_norm)
            else: 
                worksheet.write(satir_no, 0, yazarlar_metni, f_norm)
            
            worksheet.write(satir_no, 1, b['b'], f_norm)
            satir_no += 1
        satir_no += 1
    
    workbook.close()
    output.seek(0)
    return output

# --- STREAMLIT ARAYUZU ---
st.title("🖨️ IHMC 2026 - Bulut Matbaa")
st.markdown("""
Bu sistem, verileri doğrudan Google Sheets üzerinden okur.
Tüm değişiklikleri **Google Sheets** üzerinden yapın, ardından aşağıdaki butona basarak Nihai Excel çıktılarınızı indirin.
""")
st.markdown(f"**Veri Kaynağı:** [Google Sheets Tablosunu Aç]({SHEET_URL.replace('/export?format=xlsx', '')})")

if st.button("🚀 GOOGLE SHEETS'TEN VERIYI ÇEK VE MATBAAYI ÇALIŞTIR", type="primary", use_container_width=True):
    with st.spinner('☁️ Veriler buluttan indiriliyor ve işleniyor. Lütfen bekleyin...'):
        try:
            xls = pd.ExcelFile(SHEET_URL)
            
            df_ayar = pd.read_excel(xls, sheet_name=0).fillna('-')
            df_bildiriler = pd.read_excel(xls, sheet_name='Bildiriler').fillna('-')
            df_ozel = pd.read_excel(xls, sheet_name='Ozel_Etkinlikler').fillna('-')
            df_mod = pd.read_excel(xls, sheet_name='Moderatorler').fillna('-')
            
            df_bildiriler.columns = df_bildiriler.columns.str.strip()
            yazar_sutunlari = [f'Yazar {i}' for i in range(1, 9)]
            for col in yazar_sutunlari:
                if col not in df_bildiriler.columns: df_bildiriler[col] = '-'

            df_bildiriler['yazarlar'] = df_bildiriler.apply(lambda r: " - ".join([str(r[c]).strip() for c in yazar_sutunlari if str(r[c]).strip() != '-']), axis=1)
            df_bildiriler = df_bildiriler.rename(columns={'Bildiri Ismi': 'etkinlik_adi', 'Sunan Yazar': 'sunucu', 'Konu': 'konu'})
            df_bildiriler['etkinlik_adi'] = df_bildiriler['etkinlik_adi'].astype(str).str.strip()
            df_bildiriler = df_bildiriler.drop_duplicates(subset=['etkinlik_adi'])

            df_ayar['Bildiri_Adi'] = df_ayar['Bildiri_Adi'].astype(str).str.strip()
            df_master = pd.merge(df_ayar, df_bildiriler[['etkinlik_adi', 'yazarlar', 'sunucu', 'konu']], left_on='Bildiri_Adi', right_on='etkinlik_adi', how='left').fillna('-')

            df_yz = df_master[df_master['Sunum_Tipi'] == 'Yuzyuze']
            df_on = df_master[df_master['Sunum_Tipi'] == 'Online']
            
            excel_yz = excel_bas(df_yz, df_ozel, df_mod, "YUZYUZE")
            excel_on = excel_bas(df_on, df_ozel, df_mod, "ONLINE")
            
            st.success("✅ İŞLEM BAŞARILI! Nihai programlar hazır, aşağıdaki butonlardan indirebilirsiniz.")
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 YÜZYÜZE PROGRAMI İNDİR",
                    data=excel_yz,
                    file_name="KESIN_Nihai_Program_Yuzyuze.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col2:
                st.download_button(
                    label="📥 ONLINE PROGRAMI İNDİR",
                    data=excel_on,
                    file_name="KESIN_Nihai_Program_Online.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"❌ Bir hata oluştu! Lütfen Google Sheets dosyasında 'Bildiriler', 'Ozel_Etkinlikler', 'Moderatorler' sekmelerinin tam olduğundan emin olun.\n\nDetay: {e}")
