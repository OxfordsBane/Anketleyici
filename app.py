import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# Sayfa Ayarları
st.set_page_config(page_title="Hazırlık Okulu Değerlendirme Aracı", layout="wide")
st.title("🎓 İngilizce Hazırlık Değerlendirme Otomasyonu")
st.markdown("""
Bu araç, **Öğrenci Cevapları** ve **Modül Anketi** dosyalarını işleyerek 
hoca bazlı raporlar ve modül değerlendirme grafikleri oluşturur.
""")

# Dosya Yükleme Alanı
col1, col2 = st.columns(2)
with col1:
    uploaded_ogrenci = st.file_uploader("1. 'ogrenci_cevaplari.xlsx' dosyasını yükleyin", type=['xlsx', 'csv'])
with col2:
    uploaded_module = st.file_uploader("2. 'Module Evaluation Survey.xlsx' dosyasını yükleyin", type=['xlsx', 'csv'])

# Sabitler ve Ayarlar
likert_map = {
    "Strongly Agree": 5, "Agree": 4, "Neither agree, nor disagree": 3,
    "Neutral": 3, "Disagree": 2, "Strongly Disagree": 1
}
modules = [1, 2, 3, 4]

def process_files(file_ogrenci, file_module):
    # --- ZIP Oluşturmak İçin Hafıza Tamponu ---
    zip_buffer = io.BytesIO()
    
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        
        # ==========================================
        # 1. HOCA DEĞERLENDİRMELERİ İŞLEME
        # ==========================================
        try:
            df_ogrenci = pd.read_csv(file_ogrenci) if file_ogrenci.name.endswith('.csv') else pd.read_excel(file_ogrenci)
            
            # Soru Sütunlarını Belirle (İndekslere göre - Sabit yapı varsayımı)
            # Not: Dosya yapısı değişirse buradaki indeksleri (21:37) güncellemek gerekir.
            question_cols_ogrenci = df_ogrenci.columns[21:37].tolist()

            # Likert Dönüşümü
            for col in question_cols_ogrenci:
                df_ogrenci[col] = df_ogrenci[col].astype(str).str.strip().map(likert_map)

            # KEPP Ortalamaları (Okul Geneli)
            kepp_avgs = {}
            for mod in modules:
                mod_data = df_ogrenci[df_ogrenci['Modül'] == mod]
                if not mod_data.empty:
                    kepp_avgs[mod] = mod_data[question_cols_ogrenci].mean()
                else:
                    kepp_avgs[mod] = pd.Series([np.nan]*len(question_cols_ogrenci), index=question_cols_ogrenci)
            
            kepp_yearly_avg = df_ogrenci[question_cols_ogrenci].mean()

            # Her Hoca İçin Döngü
            instructors = df_ogrenci['Öğretim Elemanı'].dropna().unique()
            
            for instructor in instructors:
                clean_name = str(instructor).strip().replace('/', '-').replace('\\', '-').replace('_', ' ')
                
                # Excel'i Hafızada Oluştur
                output = io.BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}})
                workbook = writer.book
                
                # Formatlar
                header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9E1F2', 'border': 1})
                cell_fmt = workbook.add_format({'num_format': '0.00', 'align': 'center', 'border': 1})
                text_fmt = workbook.add_format({'border': 1, 'text_wrap': True})

                inst_data = df_ogrenci[df_ogrenci['Öğretim Elemanı'] == instructor]

                # --- TOTAL SHEET ---
                inst_yearly_avg = inst_data[question_cols_ogrenci].mean()
                df_total = pd.DataFrame({
                    'THE INSTRUCTOR…': question_cols_ogrenci,
                    'YOUR AVERAGE': inst_yearly_avg.values,
                    'KEPP AVERAGE': kepp_yearly_avg.values
                })
                df_total.to_excel(writer, sheet_name='TOTAL', index=False, startrow=1)
                
                worksheet = writer.sheets['TOTAL']
                worksheet.set_column('A:A', 60)
                worksheet.set_column('B:C', 15)
                
                for col_num, value in enumerate(df_total.columns.values):
                    worksheet.write(0, col_num, value, header_fmt)
                for row_num in range(len(df_total)):
                    worksheet.write(row_num + 1, 0, df_total.iloc[row_num, 0], text_fmt)
                    worksheet.write(row_num + 1, 1, df_total.iloc[row_num, 1] if pd.notna(df_total.iloc[row_num, 1]) else "-", cell_fmt)
                    worksheet.write(row_num + 1, 2, df_total.iloc[row_num, 2] if pd.notna(df_total.iloc[row_num, 2]) else "-", cell_fmt)

                # --- MOD SHEETS ---
                for mod in modules:
                    sheet_name = f'MOD {mod}'
                    inst_mod_data = inst_data[inst_data['Modül'] == mod]
                    
                    if not inst_mod_data.empty:
                        inst_mod_avg = inst_mod_data[question_cols_ogrenci].mean()
                    else:
                        inst_mod_avg = pd.Series([np.nan]*len(question_cols_ogrenci), index=question_cols_ogrenci)

                    df_mod = pd.DataFrame({
                        'THE INSTRUCTOR…': question_cols_ogrenci,
                        'YOUR AVERAGE': inst_mod_avg.values,
                        'KEPP AVERAGE': kepp_avgs[mod].values
                    })
                    df_mod.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
                    
                    worksheet = writer.sheets[sheet_name]
                    worksheet.set_column('A:A', 60)
                    worksheet.set_column('B:C', 15)
                    for col_num, value in enumerate(df_mod.columns.values):
                        worksheet.write(0, col_num, value, header_fmt)
                    for row_num in range(len(df_mod)):
                        worksheet.write(row_num + 1, 0, df_mod.iloc[row_num, 0], text_fmt)
                        worksheet.write(row_num + 1, 1, df_mod.iloc[row_num, 1] if pd.notna(df_mod.iloc[row_num, 1]) else "-", cell_fmt)
                        worksheet.write(row_num + 1, 2, df_mod.iloc[row_num, 2] if pd.notna(df_mod.iloc[row_num, 2]) else "-", cell_fmt)

                writer.close()
                output.seek(0)
                zip_file.writestr(f"Instructor_Evaluations/{clean_name}.xlsx", output.getvalue())

        except Exception as e:
            st.error(f"Öğrenci dosyası işlenirken hata oluştu: {e}")
            return None

        # ==========================================
        # 2. MODÜL ANKETİ İŞLEME
        # ==========================================
        try:
            df_module = pd.read_csv(file_module) if file_module.name.endswith('.csv') else pd.read_excel(file_module)
            question_cols_module = df_module.columns[20:27].tolist()
            level_col = 'Please choose your level. ' # CSV'de bazen sonunda boşluk olabiliyor, dikkat.

            for col in question_cols_module:
                df_module[col] = df_module[col].astype(str).str.strip().map(likert_map)
            
            mod_output = io.BytesIO()
            writer_mod = pd.ExcelWriter(mod_output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}})
            workbook_mod = writer_mod.book
            
            header_fmt_mod = workbook_mod.add_format({'bold': True, 'align': 'center', 'bg_color': '#FFE699', 'border': 1})
            cell_fmt_mod = workbook_mod.add_format({'num_format': '0.00', 'align': 'center', 'border': 1})
            text_fmt_mod = workbook_mod.add_format({'border': 1, 'text_wrap': True})

            levels = ['A1', 'A2', 'B1', 'B2']
            
            for level in levels:
                sheet_name = level
                # Level eşleşmesi bazen tam tutmayabilir, string temizliği yapalım
                df_module['clean_level'] = df_module.iloc[:, 19].astype(str).str.strip() # İndex 19 level column varsayımı
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
                    
                    # Grafik Ekleme
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
                else:
                    worksheet = workbook_mod.add_worksheet(sheet_name)
                    worksheet.write(0, 0, f"No data for Level {level}")

            writer_mod.close()
            mod_output.seek(0)
            zip_file.writestr("Module_Evaluation_Report.xlsx", mod_output.getvalue())

        except Exception as e:
            st.error(f"Modül anketi dosyası işlenirken hata oluştu: {e}")
            return None

    zip_buffer.seek(0)
    return zip_buffer

# Buton ve İşlem
if st.button("🚀 Raporları Oluştur"):
    if uploaded_ogrenci and uploaded_module:
        with st.spinner('Dosyalar işleniyor, lütfen bekleyin...'):
            result_zip = process_files(uploaded_ogrenci, uploaded_module)
            
            if result_zip:
                st.success("İşlem tamamlandı!")
                st.download_button(
                    label="📥 Tüm Raporları İndir (ZIP)",
                    data=result_zip,
                    file_name="Hazirlik_Degerlendirme_Raporlari.zip",
                    mime="application/zip"
                )
    else:
        st.warning("Lütfen her iki Excel dosyasını da yükleyin.")
