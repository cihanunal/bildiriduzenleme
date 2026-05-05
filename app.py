import streamlit as st
import pandas as pd
import xlsxwriter
import re
import datetime
import io
# Word için gerekli kütüphane
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

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
    if 'TEAMS 1' in n or 'ZOOM 1' in n or 'ODA 1' in n: return 'Teams Oda 1'
    if 'TEAMS 2' in n or 'ZOOM 2' in n or 'ODA 2' in n: return 'Teams Oda 3' # Orijinal kodda 3 yazıyordu, dokunmadım
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

# --- WORD OLUSTURMA MOTORU (YENİ) ---
def word_bas(prog, gosterilecek_ozel_etkinlikler, mod_atamalari, gun_istatistikleri, tip):
    doc = Document()
    
    # Başlık Ekleme
    title_text = "11. IHMC FİZİKİ / FACE TO FACE PROGRAM" if tip == "YUZYUZE" else "11. IHMC ONLINE (DIGITAL) PROGRAM"
    header = doc.add_heading(title_text, 0)
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ÖZEL ETKİNLİKLER
    if len(gosterilecek_ozel_etkinlikler) > 0:
        doc.add_heading("ÖZEL ETKİNLİKLER", level=1)
        for sira, ks, grup, renk_temasi, z_key in gosterilecek_ozel_etkinlikler:
            ts = str(grup.iloc[0].get('tarih_saat', '-')).strip().replace("07.05.2026", "07.05.2026 Perşembe /").replace("08.05.2026", "08.05.2026 Cuma /")
            p = doc.add_paragraph()
            p.add_run(f"{ts} | HALL: {ks}").bold = True
            
            for index, r in grup.iterrows():
                metin = "\n".join([str(r.get(c, '-')).strip() for c in ['ana_baslik', 'alt_baslik', 'sol_metin', 'sag_metin'] if str(r.get(c, '-')).strip() not in ['-', 'nan', '']])
                doc.add_paragraph(metin)
        doc.add_page_break()

    # BİLİMSEL PROGRAM
    doc.add_heading("--- BİLİMSEL BİLDİRİ PROGRAMI ---", level=1)
    
    mevcut_islenen_gun = ""
    sirali_oturumlar = sorted(prog.items(), key=lambda x: parse_zaman(x[0][0]))

    for (oturum_zaman, sal), bilds in sirali_oturumlar:
        parcalar = oturum_zaman.split(' | ')
        t = parcalar[0] if len(parcalar) > 0 else oturum_zaman
        sa = parcalar[1] if len(parcalar) > 1 else "-"
        sn = bilds[0]['sid'] if bilds else "-"

        if t != mevcut_islenen_gun:
            mevcut_islenen_gun = t
            doc.add_heading(f">>> {t.upper()} BİLİMSEL PROGRAMI <<<", level=2)
            
            # İstatistik notu
            o_sayi = gun_istatistikleri.get(t, {}).get('oturum_sayisi', 0)
            b_sayi = gun_istatistikleri.get(t, {}).get('bildiri_sayisi', 0)
            doc.add_paragraph(f"Bugün toplam {o_sayi} adet bildiri sunum oturumu gerçekleşecek ve {b_sayi} adet bildiri sunulacaktır.").italic = True

        mod_isim = mod_atamalari[(oturum_zaman, sal)]['mod']
        deg_isim = mod_atamalari[(oturum_zaman, sal)]['deg']
        
        # Oturum Bilgisi
        p_sess = doc.add_paragraph()
        p_sess.add_run(f"\n{t} | {sa} | HALL: {sal}").bold = True
        doc.add_paragraph(f"Oturum ID: {sn}\nKonu: {bilds[0]['k'].upper()}\nModerator: {mod_isim}\nDeğerlendirici: {deg_isim}")

        # Bildiri Tablosu (Metinlerin sığması için)
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Yazarlar / Authors'
        hdr_cells[1].text = 'Bildiri Adı / Title'
        
        for b in bilds:
            row_cells = table.add_row().cells
            row_cells[0].text = b['y']
            row_cells[1].text = b['b']

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- YENİ: ÖZET ÇIKARMA MOTORU ---
def oturum_ozeti_olustur(df_ayar):
    df_ozet = df_ayar[['Gun_ve_Saat', 'Salon', 'Oturum_ID', 'Sunum_Tipi']].copy()
    df_ozet = df_ozet[~df_ozet['Gun_ve_Saat'].isin(['-', 'ATANMADI', 'İPTAL EDİLDİ', 'nan', 'NaN'])]
    df_ozet = df_ozet.dropna(subset=['Gun_ve_Saat'])
    df_ozet['Salon'] = df_ozet['Salon'].apply(clean_salon)
    df_grouped = df_ozet.groupby(['Gun_ve_Saat', 'Salon', 'Oturum_ID', 'Sunum_Tipi']).size().reset_index(name='Atanan_Bildiri_Sayisi')
    df_grouped['sort_time'] = df_grouped['Gun_ve_Saat'].apply(parse_zaman)
    df_grouped = df_grouped.sort_values(by=['sort_time', 'Salon']).drop(columns=['sort_time'])
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_grouped.to_excel(writer, index=False, sheet_name='Oturum_Rontgeni')
        workbook = writer.book
        worksheet = writer.sheets['Oturum_Rontgeni']
        baslik_formati = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1})
        for col_num, value in enumerate(df_grouped.columns.values):
            worksheet.write(0, col_num, value, baslik_formati)
        worksheet.set_column('A:A', 35)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 15)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:E', 20)
    output.seek(0)
    return output

# --- ANA EXCEL OLUSTURMA MOTORU (BİLDİRİLER) - DOKUNULMADI ---
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
            m_gun_saat = str(r.get('Gun_ve_Saat', '-')).strip().replace("Persembe", "Perşembe")
            m_salon_ham = str(r.get('Salon', '-')).strip()
            m_salon = clean_salon(m_salon_ham) if m_salon_ham != '-' else '-'
            m_oturum = str(r.get('Oturum_ID', '-')).strip()
            gorev_sutunu = str(r.get('Gorev', '-')).strip().lower()
            mod_metni = f"{r.get('unvan_ad_soyad', '-')} ({r.get('kurum', '')})" if str(r.get('kurum', '-')) not in ['-', 'nan', ''] else str(r.get('unvan_ad_soyad', '-'))
            is_deg = 'deg' in gorev_sutunu
            role_key = 'deg' if is_deg else 'mod'
            hedef_musait_liste = musait_oturumlar_deg if is_deg else musait_oturumlar_mod
            hedef_serbest_liste = serbest_degs if is_deg else serbest_mods
            atandi_mi = False
            if m_gun_saat != '-' and m_salon != '-':
                hedef_tuple = (m_gun_saat, m_salon)
                if hedef_tuple in hedef_musait_liste:
                    mod_atamalari[hedef_tuple][role_key] = mod_metni
                    hedef_musait_liste.remove(hedef_tuple)
                    atandi_mi = True
            if not atandi_mi and (m_gun_saat != '-' or m_salon != '-' or m_oturum != '-'):
                for sess in hedef_musait_liste:
                    zaman_metni, salon_metni = sess[0], sess[1]
                    oturum_id_metni = prog[sess][0]['sid'] if prog[sess] else "-"
                    if (m_gun_saat == '-' or m_gun_saat in zaman_metni) and (m_salon == '-' or m_salon == salon_metni) and (m_oturum == '-' or m_oturum == oturum_id_metni):
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

    # Word Çıktısını Hazırla (Yeni Eklendi)
    word_cikti = word_bas(prog, gosterilecek_ozel_etkinlikler, mod_atamalari, gun_istatistikleri, tip)

    # Excel Yazma İşlemleri (Orijinal Kod)
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
            s_count, e_name, e_color = hucre['session_count'], hucre['event_name'], hucre.get('event_color', '#FFD966')
            if s_count > 0 and e_name:
                ws_harita.write(row_idx, col_idx, f"ÇAKIŞMA!\n[{s_count} Bildiri] VE [{e_name}]", f_map_cakisma)
                ws_harita.set_row(row_idx, 45)
            elif s_count > 0: ws_harita.write(row_idx, col_idx, f"{s_count} Bildiri Var", f_map_dolu)
            elif e_name:
                if e_color not in dinamik_renk_cache:
                    dinamik_renk_cache[e_color] = workbook.add_format({'bold': True, 'bg_color': e_color, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                ws_harita.write(row_idx, col_idx, f"[ETKİNLİK]\n{e_name}", dinamik_renk_cache[e_color])
                ws_harita.set_row(row_idx, 40)
            else: ws_harita.write(row_idx, col_idx, "BOŞ", f_map_bos)

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
        worksheet.merge_range(satir_no, 0, satir_no, 1, "ÖZEL ETKİNLİKLER", fmt['ayirici']); satir_no += 1
        for sira, ks, grup, renk_temasi, z_key in gosterilecek_ozel_etkinlikler:
            ts = str(grup.iloc[0].get('tarih_saat', '-')).strip().replace("07.05.2026", "07.05.2026 Perşembe /").replace("08.05.2026", "08.05.2026 Cuma /")
            if renk_temasi not in dinamik_renk_cache:
                dinamik_renk_cache[renk_temasi] = workbook.add_format({'bold': True, 'bg_color': renk_temasi, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
            worksheet.merge_range(satir_no, 0, satir_no, 1, f"{ts} | HALL: {ks}", dinamik_renk_cache[renk_temasi]); satir_no += 1
            for index, r in grup.iterrows():
                metin = "\n".join([str(r.get(c, '-')).strip() for c in ['ana_baslik', 'alt_baslik', 'sol_metin', 'sag_metin'] if str(r.get(c, '-')).strip() not in ['-', 'nan', '']])
                worksheet.merge_range(satir_no, 0, satir_no, 1, metin, dinamik_renk_cache[renk_temasi]); worksheet.set_row(satir_no, 50); satir_no += 1
            satir_no += 1
    worksheet.merge_range(satir_no, 0, satir_no, 1, "--- BİLİMSEL BİLDİRİ PROGRAMI ---", fmt['ayirici']); satir_no += 2
    mevcut_islenen_gun = ""
    sirali_oturumlar = sorted(prog.items(), key=lambda x: parse_zaman(x[0][0]))
    for (oturum_zaman, sal), bilds in sirali_oturumlar:
        parcalar = oturum_zaman.split(' | ')
        t, sa = parcalar[0], parcalar[1] if len(parcalar) > 1 else "-"
        sn = bilds[0]['sid'] if bilds else "-"
        if t != mevcut_islenen_gun:
            mevcut_islenen_gun = t
            worksheet.merge_range(satir_no, 0, satir_no, 1, f">>> {t.upper()} BİLİMSEL PROGRAMI <<<", fmt['mavi_gun'])
            worksheet.set_row(satir_no, 40); satir_no += 1
            worksheet.merge_range(satir_no, 0, satir_no, 1, "Her oturumdaki moderatör oturumu yönetecek...", fmt['odul_sari'])
            worksheet.set_row(satir_no, 45); satir_no += 1 
            o_sayi, b_sayi = gun_istatistikleri.get(t, {}).get('oturum_sayisi', 0), gun_istatistikleri.get(t, {}).get('bildiri_sayisi', 0)
            worksheet.merge_range(satir_no, 0, satir_no, 1, f"Bugün toplam {o_sayi} adet oturum ve {b_sayi} bildiri gerçekleşecek.", fmt['istatistik_yesil'])
            worksheet.set_row(satir_no, 25); satir_no += 2
        mod_isim, deg_isim = mod_atamalari[(oturum_zaman, sal)]['mod'], mod_atamalari[(oturum_zaman, sal)]['deg']
        worksheet.merge_range(satir_no, 0, satir_no, 1, f"{t} | {sa}", fmt['saat_baslik']); satir_no += 1
        worksheet.merge_range(satir_no, 0, satir_no, 1, f"HALL: {sal}\n{sn}\n\n{bilds[0]['k'].upper()}\nMod: {mod_isim}\nDeg: {deg_isim}", fmt['salon_baslik'])
        worksheet.set_row(satir_no, 100); satir_no += 1 
        for i, b in enumerate(bilds):
            yazarlar_metni, sunucu, is_cift = b['y'], b['s'], (i % 2 == 1) 
            f_norm, f_ul_txt = (fmt['bildiri_cift'] if is_cift else fmt['bildiri_tek']), (fmt['bildiri_cift_u_only'] if is_cift else fmt['bildiri_tek_u_only'])
            worksheet.write(satir_no, 0, yazarlar_metni, f_norm)
            worksheet.write(satir_no, 1, b['b'], f_norm)
            satir_no += 1
        satir_no += 1
    workbook.close()
    output.seek(0)
    return output, word_cikti


# --- STREAMLIT ARAYUZU ---
st.title("🖨️ IHMC 2026 - Bulut Matbaa")
st.markdown("Verileri doğrudan Google Sheets üzerinden okur. Tüm değişiklikleri oradan yapıp Nihai Excel ve Word çıktılarınızı alabilirsiniz.")

st.markdown("---")
st.subheader("1. KESİN PROGRAMI OLUŞTUR")
if st.button("🚀 MATBAAYI ÇALIŞTIR VE PROGRAMLARI ÇIKAR", type="primary", use_container_width=True):
    with st.spinner('☁️ Veriler işleniyor...'):
        try:
            xls = pd.ExcelFile(SHEET_URL)
            df_ayar = pd.read_excel(xls, sheet_name=0).fillna('-')
            df_bildiriler = pd.read_excel(xls, sheet_name='Bildiriler').fillna('-')
            df_ozel = pd.read_excel(xls, sheet_name='Ozel_Etkinlikler').fillna('-')
            df_mod = pd.read_excel(xls, sheet_name='Moderatorler').fillna('-')
            
            yazar_sutunlari = [f'Yazar {i}' for i in range(1, 9)]
            for col in yazar_sutunlari:
                if col not in df_bildiriler.columns: df_bildiriler[col] = '-'
            df_bildiriler['yazarlar'] = df_bildiriler.apply(lambda r: " - ".join([str(r[c]).strip() for c in yazar_sutunlari if str(r[c]).strip() != '-']), axis=1)
            df_bildiriler = df_bildiriler.rename(columns={'Bildiri Ismi': 'etkinlik_adi', 'Sunan Yazar': 'sunucu', 'Konu': 'konu'})
            df_ayar['Bildiri_Adi'] = df_ayar['Bildiri_Adi'].astype(str).str.strip()
            df_master = pd.merge(df_ayar, df_bildiriler[['etkinlik_adi', 'yazarlar', 'sunucu', 'konu']], left_on='Bildiri_Adi', right_on='etkinlik_adi', how='left').fillna('-')

            # Çıktıları Oluştur (YÜZYÜZE)
            df_yz = df_master[df_master['Sunum_Tipi'] == 'Yuzyuze']
            excel_yz, word_yz = excel_bas(df_yz, df_ozel, df_mod, "YUZYUZE")
            
            # Çıktıları Oluştur (ONLINE)
            df_on = df_master[df_master['Sunum_Tipi'] == 'Online']
            excel_on, word_on = excel_bas(df_on, df_ozel, df_mod, "ONLINE")
            
            st.success("✅ İŞLEM BAŞARILI!")
            
            # İNDİRME ALANI
            c1, c2 = st.columns(2)
            with c1:
                st.info("🏛️ YÜZYÜZE PROGRAMI")
                st.download_button("Excel Olarak İndir", data=excel_yz, file_name="Nihai_Program_Yuzyuze.xlsx", use_container_width=True)
                st.download_button("Word Olarak İndir (Sığdırılmış)", data=word_yz, file_name="Nihai_Program_Yuzyuze.docx", use_container_width=True)
            with c2:
                st.info("🌐 ONLINE PROGRAMI")
                st.download_button("Excel Olarak İndir", data=excel_on, file_name="Nihai_Program_Online.xlsx", use_container_width=True)
                st.download_button("Word Olarak İndir (Sığdırılmış)", data=word_on, file_name="Nihai_Program_Online.docx", use_container_width=True)

        except Exception as e:
            st.error(f"❌ Bir hata oluştu! Detay: {e}")

st.markdown("---")
st.subheader("2. OTURUM ÖZETİ")
if st.button("📊 OTURUM ÖZETİNİ ÇIKAR"):
    xls = pd.ExcelFile(SHEET_URL)
    df_ayar = pd.read_excel(xls, sheet_name=0).fillna('-')
    ozet = oturum_ozeti_olustur(df_ayar)
    st.download_button("Özeti İndir (.xlsx)", data=ozet, file_name="Oturum_Ozeti.xlsx")
