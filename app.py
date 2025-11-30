import pandas as pd
import numpy as np
import os
import shutil

# --- Configuration ---
likert_map = {
    "Strongly Agree": 5,
    "Agree": 4,
    "Neither agree, nor disagree": 3,
    "Neutral": 3,
    "Disagree": 2,
    "Strongly Disagree": 1
}

# --- 1. Process Instructor Evaluation (df_ogrenci) ---
df_ogrenci = pd.read_csv('ogrenci_cevaplari.xlsx - Worksheet.csv')
question_cols_ogrenci = df_ogrenci.columns[21:37].tolist()

# Convert Likert to Numeric
for col in question_cols_ogrenci:
    df_ogrenci[col] = df_ogrenci[col].str.strip().map(likert_map)

# Calculate KEPP Averages (School-wide averages per Module)
kepp_avgs = {}
modules = [1, 2, 3, 4]
for mod in modules:
    mod_data = df_ogrenci[df_ogrenci['Modül'] == mod]
    if not mod_data.empty:
        kepp_avgs[mod] = mod_data[question_cols_ogrenci].mean()
    else:
        kepp_avgs[mod] = pd.Series([np.nan]*len(question_cols_ogrenci), index=question_cols_ogrenci)

# Calculate KEPP Yearly Average (All data combined)
kepp_yearly_avg = df_ogrenci[question_cols_ogrenci].mean()

# Output Directory
output_dir = 'instructor_evaluations'
if os.path.exists(output_dir):
    shutil.rmtree(output_dir)
os.makedirs(output_dir)

# Process Each Instructor
instructors = df_ogrenci['Öğretim Elemanı'].dropna().unique()

for instructor in instructors:
    clean_name = str(instructor).strip().replace('/', '-').replace('\\', '-').replace('_', ' ')
    
    # Filter Instructor Data
    inst_data = df_ogrenci[df_ogrenci['Öğretim Elemanı'] == instructor]
    
    # Initialize Excel Writer with nan_inf_to_errors option to be safe
    file_path = os.path.join(output_dir, f"{clean_name}.xlsx")
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}})
    workbook = writer.book
    
    # Formats
    header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9E1F2', 'border': 1})
    cell_fmt = workbook.add_format({'num_format': '0.00', 'align': 'center', 'border': 1})
    text_fmt = workbook.add_format({'border': 1, 'text_wrap': True})
    
    # --- TOTAL Sheet ---
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
        
        # Safe Write
        val_your = df_total.iloc[row_num, 1]
        val_kepp = df_total.iloc[row_num, 2]
        
        if pd.notna(val_your):
            worksheet.write(row_num + 1, 1, val_your, cell_fmt)
        else:
            worksheet.write(row_num + 1, 1, "-", cell_fmt)
            
        if pd.notna(val_kepp):
            worksheet.write(row_num + 1, 2, val_kepp, cell_fmt)
        else:
            worksheet.write(row_num + 1, 2, "-", cell_fmt)


    # --- MOD 1-4 Sheets ---
    for mod in modules:
        sheet_name = f'MOD {mod}'
        inst_mod_data = inst_data[inst_data['Modül'] == mod]
        
        # Prepare Data
        if not inst_mod_data.empty:
            inst_mod_avg = inst_mod_data[question_cols_ogrenci].mean()
        else:
            inst_mod_avg = pd.Series([np.nan]*len(question_cols_ogrenci), index=question_cols_ogrenci)
            
        kepp_val = kepp_avgs[mod].values if not kepp_avgs[mod].empty else [np.nan]*len(question_cols_ogrenci)

        df_mod = pd.DataFrame({
            'THE INSTRUCTOR…': question_cols_ogrenci,
            'YOUR AVERAGE': inst_mod_avg.values,
            'KEPP AVERAGE': kepp_val
        })

        df_mod.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
        
        worksheet = writer.sheets[sheet_name]
        worksheet.set_column('A:A', 60)
        worksheet.set_column('B:C', 15)
        
        for col_num, value in enumerate(df_mod.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
            
        for row_num in range(len(df_mod)):
            worksheet.write(row_num + 1, 0, df_mod.iloc[row_num, 0], text_fmt)
            
            val_your = df_mod.iloc[row_num, 1]
            val_kepp = df_mod.iloc[row_num, 2]
            
            if pd.notna(val_your):
                worksheet.write(row_num + 1, 1, val_your, cell_fmt)
            else:
                 worksheet.write(row_num + 1, 1, "-", cell_fmt)
                 
            if pd.notna(val_kepp):
                worksheet.write(row_num + 1, 2, val_kepp, cell_fmt)
            else:
                worksheet.write(row_num + 1, 2, "-", cell_fmt)

    writer.close()

# --- 2. Process Module Evaluation (df_module) ---
df_module = pd.read_csv('Module Evaluation Survey.xlsx - Worksheet.csv')
question_cols_module = df_module.columns[20:27].tolist()
level_col = 'Please choose your level. '

for col in question_cols_module:
    df_module[col] = df_module[col].str.strip().map(likert_map)

module_report_path = "Module_Evaluation_Report.xlsx"
writer_mod = pd.ExcelWriter(module_report_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}})
workbook_mod = writer_mod.book

header_fmt_mod = workbook_mod.add_format({'bold': True, 'align': 'center', 'bg_color': '#FFE699', 'border': 1})
cell_fmt_mod = workbook_mod.add_format({'num_format': '0.00', 'align': 'center', 'border': 1})
text_fmt_mod = workbook_mod.add_format({'border': 1, 'text_wrap': True})

levels = ['A1', 'A2', 'B1', 'B2']

for level in levels:
    sheet_name = level
    level_data = df_module[df_module[level_col] == level]
    
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
            if pd.notna(val):
                worksheet.write(row_num + 1, 1, val, cell_fmt_mod)
            else:
                 worksheet.write(row_num + 1, 1, "-", cell_fmt_mod)
            
        chart = workbook_mod.add_chart({'type': 'column'})
        chart.add_series({
            'name':       'Average Score',
            'categories': [sheet_name, 1, 0, len(means), 0],
            'values':     [sheet_name, 1, 1, len(means), 1],
            'data_labels': {'value': True, 'num_format': '0.00'},
            'fill':       {'color': '#4472C4'}
        })
        
        chart.set_title({'name': f'{level} Level - Module Evaluation'})
        chart.set_y_axis({'name': 'Score (1-5)', 'min': 0, 'max': 5})
        chart.set_x_axis({'name': 'Questions'})
        chart.set_size({'width': 700, 'height': 400})
        
        worksheet.insert_chart('D2', chart)
        
    else:
        worksheet = workbook_mod.add_worksheet(sheet_name)
        worksheet.write(0, 0, f"No data for Level {level}")

writer_mod.close()

# Zip
final_output_dir = 'Final_Outputs'
if os.path.exists(final_output_dir):
    shutil.rmtree(final_output_dir)
os.makedirs(final_output_dir)

shutil.move('instructor_evaluations', os.path.join(final_output_dir, 'Instructor_Evaluations'))
shutil.move(module_report_path, os.path.join(final_output_dir, 'Module_Evaluation_Report.xlsx'))
shutil.make_archive('Final_Outputs', 'zip', final_output_dir)

print("Finished.")
