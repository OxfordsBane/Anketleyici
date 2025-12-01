import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import datetime

# Sayfa Ayarları
st.set_page_config(page_title="Hazırlık Okulu Değerlendirme Aracı", layout="wide")
st.title("🎓 İngilizce Hazırlık Değerlendirme Otomasyonu")
st.markdown("""
Bu araç, seçilen **Yıl** ve **Modül** kriterlerine göre verileri filtreler ve raporları oluşturur.
**Not:** "T" ile başlayan seviyeler (Örn: T1, T2) otomatik olarak değerlendirme dışı bırakılır.
""")

# --- ARAYÜZ (FİLTRELER VE DOSYA YÜKLEME) ---

st.sidebar.header("📊 Filtreleme Seçenekleri")

# 1. Yıl Seçimi
current_year = datetime.now().year
years = list(range(current_year - 1, current_year + 3)) 
selected_year = st.sidebar.selectbox("📅 Yıl Seçiniz (Anket Tarihi)", years, index=1)

# 2. Modül Seçimi
selected_module = st.sidebar.selectbox("Nx Modül Seçiniz", [1, 2, 3, 4, 5])

st.info(f"Şu an **{selected_year}** yılı **{selected_module}. Modül** verileri için rapor oluşturulacak.")

# Dosya Yükleme Alanı
col1, col2 = st.columns(2)
with col1:
    uploaded_ogrenci = st.file_uploader("1. 'ogrenci_cevaplari.xlsx' dosyasını yükleyin", type=['xlsx', 'csv'])
with col2:
    uploaded_module = st.file_uploader("2. 'Module Evaluation Survey.xlsx' dosyasını yükleyin", type=['xlsx', 'csv'])

# Sabitler
likert_map = {
    "Strongly Agree": 5, "Agree": 4, "Neither agree, nor disagree": 3,
    "Neutral": 3, "Disagree": 2, "Strongly Disagree": 1
}

def process_files(file_ogrenci, file_module, target_year, target_module):
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        
        # ==========================================
        # 1. HOCA DEĞERLENDİRMELERİ İŞLEME
        # ==========================================
        try:
            df_ogrenci = pd.read_csv(file_ogrenci) if file_ogrenci.name.endswith('.csv') else pd.read_excel(file_ogrenci)
            
            # --- FİLTRELEME ADIMLARI ---
            
            # 1. "T" ile Başlayan Seviyeleri Çıkar
            if 'Level Seviye' in df_ogrenci.columns:
                df_ogrenci = df_ogrenci[~df_ogrenci['Level Seviye'].astype(str).str.strip().str.upper().str.startswith('T')]
            
            # 2. Modül Filtresi
            df_ogrenci['Modül'] = pd.to_numeric(df_ogrenci['Modül'], errors='coerce')
            df_ogrenci = df_ogrenci[df_ogrenci['Modül'] == target_module]

            # 3. Yıl Filtresi
            if 'Tarih' in df_ogrenci.columns:
                df_ogrenci['Tarih_dt'] = pd.to_datetime(df_ogrenci['Tarih'], errors='coerce')
                df_ogrenci = df_ogrenci[df_ogrenci['Tarih_dt'].dt.year == target_year]
            
            if df_ogrenci.empty:
                st.warning(f"⚠️ Hoca Değerlendirme dosyasında kriterlere uygun veri bulunamadı! ('T' seviyeleri hariç tutuldu)")
            else:
                # --- SÜTUN BELİRLEME VE TEMİZLİK ---
                
                # Tüm olası soru sütunlarını al (21 ile 37 arası)
                raw_question_cols = df_ogrenci.columns[21:37].tolist()
                
                # ÇIKARILACAK SÜTUNLAR LİSTESİ
                # Kullanıcının istemediği veya metadata olan sütunlar
                exclude_cols = [
                    "Level Sınıf",
                    "Write your class code. (E.g. B1.01)",
                    "displays a positive and caring attitude" # Bu mükerrer veya istenmeyen bir satır
                ]
                
                # Sadece hesaplamaya katılacak soruları filtrele
                question_cols_ogrenci = [col for col in raw_question_cols if col not in exclude_cols]

                # Diğer önemli sütunlar
                comment_col = "Add any additional comments about the instructor here."
                class_col = "Write your class code. (E.g. B1.01)" # Yorumları gruplamak için kullanılacak

                # Likert Dönüşümü (Sadece seçili sorularda)
                for col in question_cols_ogrenci:
                    if col in df_ogrenci.columns:
                        df_ogrenci[col] = df_ogrenci[col].astype(str).str.strip().map(likert_map)

                # KEPP (Okul) Genel Ortalaması (Filtrelenmiş sorular üzerinden)
                kepp_avg_series = df_ogrenci[question_cols_ogrenci].mean()

                # Excel Oluşturma
                inst_output = io.BytesIO()
                writer_inst = pd.ExcelWriter(inst_output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}})
                workbook_inst = writer_inst.book
                
                # Formatlar
                header_fmt = workbook_inst.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9E1F2', 'border': 1})
                cell_fmt = workbook_inst.add_format({'num_format': '0.00', 'align': 'center', 'border': 1})
                text_fmt = workbook_inst.add_format({'border': 1, 'text_wrap': True})
                comment_main_header_fmt = workbook_inst.add_format({'bold': True, 'bg_color': '#FFEB9C', 'border': 1, 'align': 'left'})
                class_header_fmt = workbook_inst.add_format({'bold': True, 'align': 'center', 'bg_color': '#E2EFDA', 'border': 1})
                comment_text_fmt = workbook_inst.add_format({'text_wrap': True, 'border': 1, 'valign': 'top'})

                instructors = df_ogrenci['Öğretim Elemanı'].dropna().unique()

                for instructor in instructors:
                    clean_name = str(instructor).strip().replace('/', '-').replace('\\', '-').replace('_', ' ')[:31]
                    inst_data = df_ogrenci[df_ogrenci['Öğretim Elemanı'] == instructor]
                    
                    # --- PUANLAR ---
                    inst_avg_series = inst_data[question_cols_ogrenci].mean()
                    
                    df_scores = pd.DataFrame({
                        'THE INSTRUCTOR…': question_cols_ogrenci,
                        'YOUR AVERAGE': inst_avg_series.values,
                        'KEPP AVERAGE': kepp_avg_series.values
                    })

                    df_scores.to_excel(writer_inst, sheet_name=clean_name, index=False, startrow=1)
                    worksheet = writer_inst.sheets[clean_name]
                    
                    # Formatlama
                    worksheet.set_column('A:A', 60)
                    worksheet.set_column('B:C', 15)
                    for col_num, value in enumerate(df_scores.columns.values):
                        worksheet.write(0, col_num, value, header_fmt)
                    for row_num in range(len(df_scores)):
                        worksheet.write(row_num + 1, 0, df_scores.iloc[row_num, 0], text_fmt)
                        worksheet.write(row_num + 1, 1, df_scores.iloc[row_num, 1] if pd.notna(df_scores.iloc[row_num, 1]) else "-", cell_fmt)
                        worksheet.write(row_num + 1, 2, df_scores.iloc[row_num, 2] if pd.notna(df_scores.iloc[row_num, 2]) else "-", cell_fmt)

                    # --- YORUMLAR (SINIF GRUPLU) ---
                    # class_col ve comment_col varlığını kontrol et
                    if comment_col in inst_data.columns and class_col in inst_data.columns:
                        comments_df = inst_data[[class_col, comment_col]].copy()
                        comments_df = comments_df.dropna(subset=[comment_col])
                        # Sadece boşluk olan yorumları temizle
                        comments_df = comments_df[comments_df[comment_col].str.strip().astype(bool)]
                        
                        if not comments_df.empty:
                            start_row = len(df_scores) + 3
                            worksheet.write(start_row, 0, "STUDENT COMMENTS", comment_main_header_fmt)
                            current_row = start_row + 1

                            # Sınıf isimlerini temizle (NaN ise 'Unspecified' yap)
                            comments_df[class_col] = comments_df[class_col].fillna("Unspecified").astype(str).str.strip()
                            unique_classes = sorted(comments_df[class_col].unique())

                            for cls_name in unique_classes:
                                # Sınıf Başlığı
                                worksheet.merge_range(current_row, 0, current_row, 2, cls_name, class_header_fmt)
                                current_row += 1
                                
                                # O sınıfa ait yorumlar
                                cls_comments = comments_df[comments_df[class_col] == cls_name][comment_col].tolist()
                                for comment in cls_comments:
                                    worksheet.write(current_row, 0, str(comment).strip(), comment_text_fmt)
                                    current_row += 1

                writer_inst.close()
                inst_output.seek(0)
                zip_file.writestr("Instructor_Evaluations.xlsx", inst_output.getvalue())

        except Exception as e:
            st.error(f"Hoca değerlendirme dosyası işlenirken hata: {e}")
            return None

        # ==========================================
        # 2. MODÜL ANKETİ İŞLEME
        # ==========================================
        try:
            df_module = pd.read_csv(file_module) if file_module.name.endswith('.csv') else pd.read_excel(file_module)
            
            # --- FİLTRELEME ADIMI ---
            df_module['Modül'] = pd.to_numeric(df_module['Modül'], errors='coerce')
            df_module = df_module[df_module['Modül'] == target_module]

            if df_module.empty:
                st.warning(f"⚠️ Modül Değerlendirme dosyasında {target_module}. Modül için veri bulunamadı!")
            else:
                # Sütunları Belirle (20-27 arası)
                # İndeks 26 genelde "Add your comments..." oluyor, onu ortalamadan çıkaracağız.
                # Sadece ilk 6 soru puanlanabilir soru (Index 20-25 arası)
                question_cols_module = df_module.columns[20:26].tolist()
                
                # Yorum sütununu isminden bulmaya çalışalım (genelde sonlarda olur)
                # "Add your comments" içeren sütunu bul
                comment_col_mod = [c for c in df_module.columns if "Add your comments" in str(c)]
                if comment_col_mod:
                    comment_col_mod = comment_col_mod[0] # İlk eşleşeni al
                else:
                    comment_col_mod = None

                # Likert Dönüşümü
                for col in question_cols_module:
                    df_module[col] = df_module[col].astype(str).str.strip().map(likert_map)
                
                mod_output = io.BytesIO()
                writer_mod = pd.ExcelWriter(mod_output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}})
                workbook_mod = writer_mod.book
                
                header_fmt_mod = workbook_mod.add_format({'bold': True, 'align': 'center', 'bg_color': '#FFE699', 'border': 1})
                cell_fmt_mod = workbook_mod.add_format({'num_format': '0.00', 'align': 'center', 'border': 1})
                text_fmt_mod = workbook_mod.add_format({'border': 1, 'text_wrap': True})
                comment_header_mod = workbook_mod.add_format({'bold': True, 'bg_color': '#BDD7EE', 'border': 1})

                levels = ['A1', 'A2', 'B1', 'B2']
                
                for level in levels:
                    sheet_name = level
                    df_module['clean_level'] = df_module.iloc[:, 19].astype(str).str.strip()
                    level_data = df_module[df_module['clean_level'] == level]

                    if not level_data.empty:
                        # 1. ORTALAMALAR TABLOSU
                        means = level_data[question_cols_module].mean().reset_index()
                        means.columns = ['Question', 'Average Score']
                        means.to_excel(writer_mod, sheet_name=sheet_name, index=False, startrow=1)
                        
                        worksheet = writer_mod.sheets[sheet_name]
                        worksheet.set_column('A:A', 70)
                        worksheet.set_column('B:B', 15)
                        worksheet.write(0, 0, 'Question', header_fmt_mod)
                        worksheet.write(0, 1, 'Average Score', header_fmt_mod)
                        
                        for row_num in range(len(means)):
                            worksheet.write(row_num + 1, 0, means.iloc[row_num, 0], text_fmt_mod)
                            val = means.iloc[row_num, 1]
                            worksheet.write(row_num + 1, 1, val if pd.notna(val) else "-", cell_fmt_mod)
                        
                        # 2. GRAFİK
                        chart = workbook_mod.add_chart({'type': 'column'})
                        chart.add_series({
                            'name': 'Average Score',
                            'categories': [sheet_name, 1, 0, len(means), 0],
                            'values': [sheet_name, 1, 1, len(means), 1],
                            'data_labels': {'value': True, 'num_format': '0.00'},
                            'fill': {'color': '#4472C4'}
                        })
                        chart.set_title({'name': f'{level} Level - Module Evaluation'})
                        chart.set_y_axis({'name': 'Score (1-5)', 'min': 0, 'max': 5})
                        chart.set_size({'width': 700, 'height': 400})
                        worksheet.insert_chart('D2', chart)

                        # 3. YORUMLAR (TABLONUN VE GRAFİĞİN ALTINA)
                        if comment_col_mod:
                            mod_comments = level_data[comment_col_mod].dropna().astype(str).tolist()
                            mod_comments = [c for c in mod_comments if c.strip()] # Boşları at
                            
                            if mod_comments:
                                comment_start_row = len(means) + 25 # Grafiğin altına denk gelmesi için boşluk
                                worksheet.write(comment_start_row, 0, "STUDENT COMMENTS", comment_header_mod)
                                
                                for idx, com in enumerate(mod_comments):
                                    worksheet.write(comment_start_row + 1 + idx, 0, com, text_fmt_mod)

                    else:
                        worksheet = workbook_mod.add_worksheet(sheet_name)
                        worksheet.write(0, 0, f"No data for Level {level}")

                writer_mod.close()
                mod_output.seek(0)
                zip_file.writestr("Module_Evaluation_Report.xlsx", mod_output.getvalue())

        except Exception as e:
            st.error(f"Modül anketi dosyası işlenirken hata: {e}")
            return None

    zip_buffer.seek(0)
    return zip_buffer

# Buton ve İşlem
if st.button("🚀 Raporları Oluştur"):
    if uploaded_ogrenci and uploaded_module:
        with st.spinner('Dosyalar işleniyor, lütfen bekleyin...'):
            result_zip = process_files(uploaded_ogrenci, uploaded_module, selected_year, selected_module)
            
            if result_zip:
                st.success(f"İşlem tamamlandı! {selected_year} - Modül {selected_module} raporları hazır.")
                st.download_button(
                    label="📥 Raporları İndir (ZIP)",
                    data=result_zip,
                    file_name=f"Hazirlik_Raporlari_{selected_year}_Modul{selected_module}.zip",
                    mime="application/zip"
                )
    else:
        st.warning("Lütfen her iki Excel dosyasını da yükleyin.")
