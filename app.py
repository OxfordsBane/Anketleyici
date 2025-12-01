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

# --- HEDEF SORU LİSTESİ (SABİT) ---
TARGET_QUESTIONS = [
    "comes prepared with materials to be used in lessons.",
    "starts and ends lessons on time.",
    "teaches the course content clearly.",
    "speaks English clearly and comprehensibly.",
    "has an attitude that supports student learning outside the classroom.",
    "encourages students to participate in class.",
    "keeps a regular record of student attendance and timeliness.",
    "uses class time efficiently and effectively.",
    "uses office hours efficiently and fairly.",
    "has adapted to technological advancements.",
    "enters and announces the necessary records. (Attendance, grades, scores, etc.)",
    "doesn't speak Turkish in class unless necessary.",
    "is good at classroom management.",
    "displays a positive and caring attitude.",
    "has good overall performance.",
    "creates a motivating and convenient learning environment in class."
]

# --- ARAYÜZ ---
st.sidebar.header("📊 Filtreleme Seçenekleri")

current_year = datetime.now().year
years = list(range(current_year - 1, current_year + 3)) 
selected_year = st.sidebar.selectbox("📅 Yıl Seçiniz (Anket Tarihi)", years, index=1)
selected_module = st.sidebar.selectbox("Nx Modül Seçiniz", [1, 2, 3, 4, 5])

st.info(f"Şu an **{selected_year}** yılı **{selected_module}. Modül** verileri için rapor oluşturulacak.")

col1, col2 = st.columns(2)
with col1:
    uploaded_ogrenci = st.file_uploader("1. 'ogrenci_cevaplari.xlsx' dosyasını yükleyin", type=['xlsx', 'csv'])
with col2:
    uploaded_module = st.file_uploader("2. 'Module Evaluation Survey.xlsx' dosyasını yükleyin", type=['xlsx', 'csv'])

likert_map = {
    "Strongly Agree": 5, "Agree": 4, "Neither agree, nor disagree": 3,
    "Neutral": 3, "Disagree": 2, "Strongly Disagree": 1
}

def clean_column_names(df):
    # Sadece çift tırnakları ve gereksiz boşlukları temizle (tek tırnak kalmalı)
    df.columns = df.columns.str.strip().str.replace('"', '')
    return df

def process_files(file_ogrenci, file_module, target_year, target_module):
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        
        # ==========================================
        # 1. HOCA DEĞERLENDİRMELERİ İŞLEME
        # ==========================================
        try:
            df_ogrenci = pd.read_csv(file_ogrenci) if file_ogrenci.name.endswith('.csv') else pd.read_excel(file_ogrenci)
            df_ogrenci = clean_column_names(df_ogrenci)
            
            # --- FİLTRELEME ---
            if 'Level Seviye' in df_ogrenci.columns:
                df_ogrenci = df_ogrenci[~df_ogrenci['Level Seviye'].astype(str).str.strip().str.upper().str.startswith('T')]
            
            df_ogrenci['Modül'] = pd.to_numeric(df_ogrenci['Modül'], errors='coerce')
            df_ogrenci = df_ogrenci[df_ogrenci['Modül'] == target_module]

            if 'Tarih' in df_ogrenci.columns:
                df_ogrenci['Tarih_dt'] = pd.to_datetime(df_ogrenci['Tarih'], errors='coerce')
                df_ogrenci = df_ogrenci[df_ogrenci['Tarih_dt'].dt.year == target_year]
            
            if df_ogrenci.empty:
                st.warning(f"⚠️ Hoca Değerlendirme dosyasında kriterlere uygun veri bulunamadı!")
            else:
                # Soru Sütunlarını Belirle
                available_questions = []
                seen = set()
                for q in TARGET_QUESTIONS:
                    if q in df_ogrenci.columns and q not in seen:
                        available_questions.append(q)
                        seen.add(q)
                
                question_cols_ogrenci = available_questions
                comment_col = "Add any additional comments about the instructor here."

                # Likert Dönüşümü
                for col in question_cols_ogrenci:
                    df_ogrenci[col] = df_ogrenci[col].astype(str).str.strip().map(likert_map)

                # KEPP Ortalaması
                kepp_avg_series = df_ogrenci[question_cols_ogrenci].mean()
                
                # Sınıf Adı Oluşturma
                if 'Level Seviye' in df_ogrenci.columns and 'Level Sınıf' in df_ogrenci.columns:
                    s_seviye = df_ogrenci['Level Seviye'].astype(str).str.strip()
                    s_sinif = df_ogrenci['Level Sınıf'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                    s_sinif = s_sinif.apply(lambda x: x.zfill(2) if x.isdigit() else x)
                    df_ogrenci['Calculated_Class_Code'] = s_seviye + "." + s_sinif
                    class_col = 'Calculated_Class_Code'
                else:
                    class_col = "Write your class code. (E.g. B1.01)"

                # Excel Başlat
                inst_output = io.BytesIO()
                writer_inst = pd.ExcelWriter(inst_output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}})
                workbook_inst = writer_inst.book
                
                # --- FORMATLAR ---
                header_fmt = workbook_inst.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9E1F2', 'border': 1})
                cell_fmt = workbook_inst.add_format({'num_format': '0.00', 'align': 'center', 'border': 1})
                text_fmt = workbook_inst.add_format({'border': 1, 'text_wrap': True})
                comment_main_header_fmt = workbook_inst.add_format({'bold': True, 'bg_color': '#FFEB9C', 'border': 1, 'align': 'left'})
                comment_text_fmt = workbook_inst.add_format({'text_wrap': True, 'border': 1, 'valign': 'top'})

                # RENKLİ SEVİYE FORMATLARI (Yazı: Beyaz, Bold, Center)
                fmt_a1 = workbook_inst.add_format({'bg_color': '#F5BD02', 'font_color': 'white', 'bold': True, 'align': 'center', 'border': 1})
                fmt_a2 = workbook_inst.add_format({'bg_color': '#F07F09', 'font_color': 'white', 'bold': True, 'align': 'center', 'border': 1})
                fmt_b1 = workbook_inst.add_format({'bg_color': '#9F2936', 'font_color': 'white', 'bold': True, 'align': 'center', 'border': 1})
                fmt_b2 = workbook_inst.add_format({'bg_color': '#4E8542', 'font_color': 'white', 'bold': True, 'align': 'center', 'border': 1})
                fmt_default = workbook_inst.add_format({'bg_color': '#E2EFDA', 'bold': True, 'align': 'center', 'border': 1})

                instructors = df_ogrenci['Öğretim Elemanı'].dropna().unique()

                for instructor in instructors:
                    clean_name = str(instructor).strip().replace('/', '-').replace('\\', '-').replace('_', ' ')[:31]
                    inst_data = df_ogrenci[df_ogrenci['Öğretim Elemanı'] == instructor]
                    
                    inst_avg_series = inst_data[question_cols_ogrenci].mean()
                    
                    df_scores = pd.DataFrame({
                        'THE INSTRUCTOR…': question_cols_ogrenci,
                        'YOUR AVERAGE': inst_avg_series.values,
                        'KEPP AVERAGE': kepp_avg_series.values
                    })

                    df_scores.to_excel(writer_inst, sheet_name=clean_name, index=False, startrow=1)
                    worksheet = writer_inst.sheets[clean_name]
                    
                    worksheet.set_column('A:A', 60)
                    worksheet.set_column('B:C', 15)
                    for col_num, value in enumerate(df_scores.columns.values):
                        worksheet.write(0, col_num, value, header_fmt)
                    for row_num in range(len(df_scores)):
                        worksheet.write(row_num + 1, 0, df_scores.iloc[row_num, 0], text_fmt)
                        worksheet.write(row_num + 1, 1, df_scores.iloc[row_num, 1] if pd.notna(df_scores.iloc[row_num, 1]) else "-", cell_fmt)
                        worksheet.write(row_num + 1, 2, df_scores.iloc[row_num, 2] if pd.notna(df_scores.iloc[row_num, 2]) else "-", cell_fmt)

                    # Yorumlar
                    if comment_col in inst_data.columns and class_col in inst_data.columns:
                        comments_df = inst_data[[class_col, comment_col]].copy()
                        comments_df = comments_df.dropna(subset=[comment_col])
                        comments_df = comments_df[comments_df[comment_col].str.strip().astype(bool)]
                        
                        if not comments_df.empty:
                            start_row = len(df_scores) + 3
                            worksheet.write(start_row, 0, "STUDENT COMMENTS", comment_main_header_fmt)
                            current_row = start_row + 1

                            comments_df[class_col] = comments_df[class_col].fillna("Unspecified").astype(str).str.strip()
                            unique_classes = sorted(comments_df[class_col].unique())

                            for cls_name in unique_classes:
                                # Seviyeyi belirle (A1, A2...)
                                level_prefix = cls_name.split('.')[0].upper()
                                
                                # Rengi Seç
                                if level_prefix == 'A1':
                                    current_fmt = fmt_a1
                                elif level_prefix == 'A2':
                                    current_fmt = fmt_a2
                                elif level_prefix == 'B1':
                                    current_fmt = fmt_b1
                                elif level_prefix == 'B2':
                                    current_fmt = fmt_b2
                                else:
                                    current_fmt = fmt_default

                                # Sınıf Başlığını Yaz (Sadece A Sütunu)
                                worksheet.write(current_row, 0, cls_name, current_fmt)
                                current_row += 1
                                
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
            df_module = clean_column_names(df_module)

            # --- FİLTRELEME ---
            df_module['Modül'] = pd.to_numeric(df_module['Modül'], errors='coerce')
            df_module = df_module[df_module['Modül'] == target_module]

            if df_module.empty:
                st.warning(f"⚠️ Modül Değerlendirme dosyasında {target_module}. Modül için veri bulunamadı!")
            else:
                question_cols_module = df_module.columns[20:26].tolist()
                comment_col_mod = [c for c in df_module.columns if "Add your comments" in str(c)]
                comment_col_mod = comment_col_mod[0] if comment_col_mod else None

                for col in question_cols_module:
                    df_module[col] = df_module[col].astype(str).str.strip().map(likert_map)
                
                mod_output = io.BytesIO()
                writer_mod = pd.ExcelWriter(mod_output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}})
                workbook_mod = writer_mod.book
                
                header_fmt_mod = workbook_mod.add_format({'bold': True, 'align': 'center', 'bg_color': '#FFE699', 'border': 1})
                cell_fmt_mod = workbook_mod.add_format({'num_format': '0.00', 'align': 'center', 'border': 1})
                text_fmt_mod = workbook_mod.add_format({'border': 1, 'text_wrap': True})
                comment_header_mod = workbook_mod.add_format({'bold': True, 'bg_color': '#BDD7EE', 'border': 1})

                # --- 1. OVERALL SHEET ---
                sheet_name = "OVERALL"
                means_total = df_module[question_cols_module].mean().reset_index()
                means_total.columns = ['Question', 'Average Score']
                means_total.to_excel(writer_mod, sheet_name=sheet_name, index=False, startrow=1)
                
                worksheet = writer_mod.sheets[sheet_name]
                worksheet.set_column('A:A', 70)
                worksheet.set_column('B:B', 15)
                worksheet.write(0, 0, 'Question', header_fmt_mod)
                worksheet.write(0, 1, 'Average Score', header_fmt_mod)
                
                for row_num in range(len(means_total)):
                    worksheet.write(row_num + 1, 0, means_total.iloc[row_num, 0], text_fmt_mod)
                    val = means_total.iloc[row_num, 1]
                    worksheet.write(row_num + 1, 1, val if pd.notna(val) else "-", cell_fmt_mod)
                
                chart = workbook_mod.add_chart({'type': 'column'})
                chart.add_series({
                    'name': 'Average Score',
                    'categories': [sheet_name, 1, 0, len(means_total), 0],
                    'values': [sheet_name, 1, 1, len(means_total), 1],
                    'data_labels': {'value': True, 'num_format': '0.00'},
                    'fill': {'color': '#4472C4'}
                })
                chart.set_title({'name': f'OVERALL (All Levels) - Module Evaluation'})
                chart.set_y_axis({'name': 'Score (1-5)', 'min': 0, 'max': 5})
                chart.set_size({'width': 700, 'height': 400})
                worksheet.insert_chart('D2', chart)
                
                if comment_col_mod:
                    all_comments = df_module[comment_col_mod].dropna().astype(str).tolist()
                    all_comments = [c for c in all_comments if c.strip()]
                    if all_comments:
                        comment_start_row = len(means_total) + 25 
                        worksheet.write(comment_start_row, 0, "STUDENT COMMENTS (ALL LEVELS)", comment_header_mod)
                        for idx, com in enumerate(all_comments):
                            worksheet.write(comment_start_row + 1 + idx, 0, com, text_fmt_mod)

                # --- 2. LEVEL SHEETS ---
                levels = ['A1', 'A2', 'B1', 'B2']
                for level in levels:
                    sheet_name = level
                    df_module['clean_level'] = df_module.iloc[:, 19].astype(str).str.strip()
                    level_data = df_module[df_module['clean_level'] == level]

                    if not level_data.empty:
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

                        if comment_col_mod:
                            mod_comments = level_data[comment_col_mod].dropna().astype(str).tolist()
                            mod_comments = [c for c in mod_comments if c.strip()]
                            
                            if mod_comments:
                                comment_start_row = len(means) + 25 
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
