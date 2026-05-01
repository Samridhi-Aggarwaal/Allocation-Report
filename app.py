# -------------------------
# Imports & Constants
# -------------------------
import os
import re
import sys
import difflib
import calendar
import tempfile
import pandas as pd
from datetime import datetime
import win32com.client as win32
from contextlib import suppress
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QRunnable, QThreadPool, pyqtSignal, QObject
from PyQt5.QtWidgets import QMessageBox, QFileDialog, QHBoxLayout, QDateEdit

PALETTE = [    
  "#ede7e3", "#bbdefb", "#f8bbd0", "#c8e6c9", "#fff9c4",
  "#e1bee7", "#d1c4e9", "#ffccbc", "#b2ebf2", "#cfd8dc"
]
APAC_CLIENT_NAME = "APAC - Central Cost"
COMPANY_RECIPIENTS = {
  "company": ["abc@xyz.com"]
}
UNIVERSAL_RECIPIENTS = ["abc@xyz.com"]

PRIORITY_ORDER = ['Billed', 'PreSow', 'Blocked Code', 'Client Bench', 'Fractal Bench', 'Cluster Bench', 'Investment']
SOW_CAT_TYPE = pd.CategoricalDtype(categories=PRIORITY_ORDER, ordered=True)

GUIDELINES_COLUMNS = [
  {"Sheet Name": "Client Allocation Summary", "Details": "Displays the number of employees working for each client under each Statement of Work (SOW). The count is determined by adding the allocation percentage for each SOW per client and then dividing by 100."},
  {"Sheet Name": "Monthly Headcount Summary", "Details": "Displays the headcount of employees working for each client under each Statement of Work (SOW) in the entire month. The count is determined by adding the number of employees working for each client under each SOW."},
  {"Sheet Name": "Total Headcount Summary", "Details": "Displays the headcount of employees working for each client under each Statement of Work (SOW) for today. The count is determined by adding the number of employees working for each client under each SOW as of today."},
  {"Sheet Name": "Allocation Comparison", "Details": "Displays a list of a comparison of today’s allocation data with data from the previous working day and the previous week, highlighting the differences in allocation."},
  {"Sheet Name": "Billed Employee Details", "Details": "Displays a list of all billable employees, organized by client."},
  {"Sheet Name": "Unbilled Employee Details", "Details": "Displays a list of all non-billable employees, organized by client."},
  {"Sheet Name": "Allocations Ending This Month", "Details": "Displays a list of all employees whose projects are concluding within the current month."}
]

TODAY_COL_TEMPLATE = "Count for Today ({})"
PREV_COL_TEMPLATE = "Count for Previous day ({})"
WEEK_COL_TEMPLATE = "Count for Previous week ({})"
CHANGE_PREV_TEMPLATE = "Change from Previous day ({})"
CHANGE_WEEK_TEMPLATE = "Change from Previous week ({})"

# -------------------------
# Helper functions
# -------------------------
def get_count_on_date(df, date):
  return df[df['Date'] == date].groupby(['Client', 'Sow'], as_index=False)['No. of Employees'].sum()
  
def validate_columns(df, required_columns):
  missing = [col for col in required_columns if col not in df.columns]
  if missing:
    raise ValueError(f"Missing columns in input file: {missing}")

def normalize_company_name(name):
  if not isinstance(name, str):
    return ""
  return re.sub(r'[^a-z0-9]', '', name.lower())

def parse_dates(df):
  if 'Date' not in df.columns:
    raise ValueError("DataFrame must contain a 'Date' column.")
  df['Date'] = pd.to_datetime(df['Date'], format="%Y-%m-%d", errors='coerce')
  if df['Date'].isnull().all():
    df['Date'] = pd.to_datetime(df['Date'], format="%d/%m/%Y %I:%M:%S %p", errors='coerce')
  elif 'Date' in df.columns:
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
  else:
    df['Date'] = pd.to_datetime('today').normalize()
  df['Date'] = df['Date'].dt.normalize()
  return df

def df_to_html_with_row_color(df, color_col, apply_color=True):
  if df is None or df.empty:
    return "<i>No data available</i>"
  color_map = {}
  if apply_color and color_col in df.columns:
    color_map = {val: PALETTE[i % len(PALETTE)] for i, val in enumerate(df[color_col].unique())}
  html_parts = ['<table border="1" cellpadding="4" style="border-collapse:collapse; font-size:10px; text-align:center;">']
  html_parts.append('<tr>' + ''.join(f'<th>{col}</th>' for col in df.columns) + '</tr>')
  for _, row in df.iterrows():
    bg_style = f'background-color:{color_map.get(row.get(color_col), "")};' if apply_color else ''
    row_html = ''.join(
      f'<td style="{bg_style} text-align:center;">{row[col]}</td>'
      for col in df.columns
    )
    html_parts.append(f'<tr>{row_html}</tr>')
    html_parts.append('</table>')
  return ''.join(html_parts)

def normalize_sow(sow_value):
  sow_str = str(sow_value).strip().lower()
  if 'blocked' in sow_str:
    return 'Blocked Code'
  elif 'client' in sow_str:
    return 'Client Bench'
  elif 'fractal' in sow_str:
    return 'Fractal Bench'
  elif 'sea' in sow_str:
    return 'Cluster Bench'
  elif 'presow' in sow_str or 'pre-sow' in sow_str:
    return 'PreSow'
  elif 'investment' in sow_str:
    return 'Investment'
  elif 'upcoming' in sow_str:
    return 'Upcoming Project'
  elif 'core delivery' in sow_str:
    return 'Core Delivery'
  else:
    return 'Billed'
    
def add_skill_columns(df, file_path):
  csv_path = os.path.join(os.path.dirname(file_path), "skills.csv")
  csv_path = os.path.normpath(csv_path)
  if not os.path.exists(csv_path):
    print(f"CSV file not found at {csv_path}. Continuing without skills data.")
    return df
  try:
    skills_df = pd.read_csv(csv_path)
    if 'Employee ID' in skills_df.columns:
      skills_df.rename(columns={'Employee ID': 'Employee Code'}, inplace=True)
    required_cols = ['Employee Code', 'Skill Item', 'Skill Item Category', 'Skill Item Category Group']
    missing_cols = [col for col in required_cols if col not in skills_df.columns]
    if missing_cols:
      print(f"Missing required columns in skills.csv: {missing_cols}. Skipping merge, continuing without skills data.")
      return df
    skills_df[required_cols] = skills_df[required_cols].fillna('')
    aggregated_skills = skills_df.groupby('Employee Code').agg({
      'Skill Item': lambda x: ', '.join(sorted(filter(None, set(x)))),
      'Skill Item Category': lambda x: ', '.join(sorted(filter(None, set(x)))),
      'Skill Item Category Group': lambda x: ', '.join(sorted(filter(None, set(x))))
    }).reset_index()
    print("\nAll required columns are present. Proceeding with merging skill data into main DataFrame...")
    merged_df = df.merge(aggregated_skills, on='Employee Code', how='left')
    output_path = os.path.join(os.path.dirname(file_path), "Merged_df.xlsx")
    output_path = os.path.normpath(output_path)
    return merged_df
  except Exception as e:
    print(f"\nError processing skills.csv: {e}. Continuing without skills data.")
    return df
def map_column_names(df, expected_cols):
  if df is None:
    return df
  actual_norm_map = {}
  for c in df.columns:
    norm = re.sub(r'[^a-z0-9]', '', str(c).lower())
    actual_norm_map.setdefault(norm, c)
  expected_norm_map = {}
  for exp in expected_cols:
    norm = re.sub(r'[^a-z0-9]', '', str(exp).lower())
    expected_norm_map.setdefault(norm, exp)
  rename_map = {}
    
  for actual_col in list(df.columns):
    actual_norm = re.sub(r'[^a-z0-9]', '', str(actual_col).lower())
    if actual_norm in expected_norm_map:
      target_expected = expected_norm_map[actual_norm]
      if actual_col == target_expected:
        continue
      if target_expected in df.columns:
        continue
      rename_map[actual_col] = target_expected
    else:
      matches = difflib.get_close_matches(actual_norm, list(expected_norm_map.keys()), n=1, cutoff=0.85)
      if matches:
        target_expected = expected_norm_map[matches[0]]
        if target_expected in df.columns:
          continue
        rename_map[actual_col] = target_expected
      
    if rename_map:
      try:
        df = df.rename(columns=rename_map)
      except Exception:
        return df
    return df 

# -------------------------
# Data processing functions
# -------------------------
def update_allocation_history(merged_table, history_file, selected_date):
  try:
    if merged_table is None or merged_table.empty:
      print("Merged table is empty or None. Skipping history update.")
      return
    merged_table = merged_table.copy()
    merged_table['Date'] = selected_date
    merged_table = merged_table[['Client', 'Sow', 'No. of Employees', 'Date']]
    merged_table['Sow'] = merged_table['Sow'].astype(SOW_CAT_TYPE)
    if os.path.exists(history_file):
      try:
        history_df = pd.read_excel(history_file, sheet_name='History', engine='openpyxl')
        print("Existing history file loaded.")
      except Exception as read_error:
        print(f"Failed to read existing history file: {read_error}")
        history_df = pd.DataFrame(columns=['Client', 'Sow', 'No. of Employees', 'Date'])
    else:
      print("History file not found. Creating new history file.")
      history_df = pd.DataFrame(columns=['Client', 'Sow', 'No. of Employees', 'Date'])
    
    history_df['Date'] = pd.to_datetime(history_df['Date'], errors='coerce').dt.date
    merged_table['is_duplicate'] = merged_table.apply(
      lambda row: (
        (history_df['Client'] == row['Client']) &
        (history_df['Sow'] == row['Sow']) &
        (history_df['No. of Employees'] == row['No. of Employees']) &
        (history_df['Date'] == row['Date'])
      ).any(),
      axis=1
    )
    
    new_rows = merged_table[~merged_table['is_duplicate']].drop(columns='is_duplicate')
    if new_rows.empty:
      print("All rows for today already exist in history. Skipping save operation.")
      return
    else:
      print(f"Appending {len(new_rows)} new rows to history.")
      history_df = pd.concat([history_df, new_rows], ignore_index=True)
   
    before_dedup = len(history_df)
    history_df.drop_duplicates(subset=['Client', 'Sow', 'No. of Employees', 'Date'], keep='last', inplace=True)
    after_dedup = len(history_df)
    removed = before_dedup - after_dedup
    if removed > 0:
      print(f"Removed {removed} duplicate rows from history.")
    history_df['Sow'] = history_df['Sow'].astype(SOW_CAT_TYPE)
    history_df.sort_values(by=['Date', 'Client', 'Sow'], inplace=True)

    with pd.ExcelWriter(history_file, engine='xlsxwriter') as writer:
      history_df.to_excel(writer, sheet_name='History', index=False)
      worksheet = writer.sheets['History']
      cell_format = writer.book.add_format({'align': 'center'})                                                                    # pyright: ignore[reportAttributeAccessIssue]
      
      for idx, col in enumerate(history_df.columns):
        max_len = max(history_df[col].astype(str).map(len).max(), len(col) + 2)
        worksheet.set_column(idx, idx, max_len, cell_format)
    
    except Exception as main_error:
      print(f"Error updating history file: {main_error}")
      raise

def generate_employee_change_summary(input_path, merged_table, start_date="2025-01-01", days_back=7):
  df = pd.read_excel(input_path)
  df = map_column_names(df, ['Date', 'Client', 'Sow', 'No. of Employees', 'Start Date', 'End Date'])
  required_columns = ['Date', 'Client', 'Sow', 'No. of Employees']
  validate_columns(df, required_columns)
  
  df = parse_dates(df)
  
  today = pd.to_datetime("today").normalize()
  seven_days_ago = today - pd.Timedelta(days=days_back)
  start_date = pd.to_datetime(start_date)
  today_col = TODAY_COL_TEMPLATE.format(today.date())
  today_df = get_count_on_date(df, today).rename(columns={'No. of Employees': today_col})                                                                    # pyright: ignore[reportCallIssue]
  clients_for_combination = merged_table[['Client']].drop_duplicates().copy()
  
  if 'APAC_CLIENT_NAME' in globals() and APAC_CLIENT_NAME not in clients_for_combination['Client'].values:
    clients_for_combination = pd.concat([clients_for_combination, pd.DataFrame({'Client': [APAC_CLIENT_NAME]})], ignore_index=True)
  all_combinations_full = clients_for_combination.merge(pd.DataFrame({'Sow': PRIORITY_ORDER}), how='cross')
  merged = all_combinations_full.merge(today_df, on=['Client', 'Sow'], how='left')
  
  recent_dates = df[(df['Date'] < today) & (df['Date'] > seven_days_ago)]['Date'].unique()
  if len(recent_dates) > 0:
    closest_prev_date = max(recent_dates)
    prev_col = PREV_COL_TEMPLATE.format(closest_prev_date.date())
    prev_df = get_count_on_date(df, closest_prev_date).rename(columns={'No. of Employees': prev_col})
    merged = merged.merge(prev_df, on=['Client', 'Sow'], how='left').fillna(0)
    change_prev_col = CHANGE_PREV_TEMPLATE.format(closest_prev_date.date())
    merged[change_prev_col] = merged[today_col] - merged[prev_col]

  fixed_past_range = df[(df['Date'] <= seven_days_ago) & (df['Date'] >= start_date)]['Date'].unique()
  if len(fixed_past_range) > 0:
    fixed_past_date = max(fixed_past_range)
    fixed_past_col = WEEK_COL_TEMPLATE.format(fixed_past_date.date())
    fixed_past_df = get_count_on_date(df, fixed_past_date).rename(columns={'No. of Employees': fixed_past_col})
    merged = merged.merge(fixed_past_df, on=['Client', 'Sow'], how='left').fillna(0)
    change_fixed_col = CHANGE_WEEK_TEMPLATE.format(fixed_past_date.date())
    merged[change_fixed_col] = merged[today_col] - merged[fixed_past_col]
    
  final_df = merged.sort_values(by=['Client', 'Sow'], ignore_index=True)
  return final_df

# -------------------------
# Worker / Signals
# -------------------------
class WorkerSignals(QObject):
  finished = pyqtSignal(object, object, object)
class TableWorker(QRunnable):
  def __init__(self, file_path, generate_func):
    super().__init__()
    self.file_path = file_path
    self.generate_func = generate_func
    self.signals = WorkerSignals()
    
  def run(self):
    html, dfs, error = self.generate_func(self.file_path)
    self.signals.finished.emit(html, dfs, error)

# -------------------------
# Main Application
# -------------------------
class ExcelPivotApp(QtWidgets.QWidget):
  def __init__(self):
    super().__init__()
    self.pivot_df = None
    self.last_file_path = None
    self.last_file_mtime = None
    self.threadpool = QThreadPool()
    self.selected_date = datetime.today().date()
    self.init_ui()
  
  def init_ui(self):
    self.setWindowTitle('Excel Pivot Table Generator')
    self.setWindowIcon(QtGui.QIcon('excel-pivot-ui-app/icon.png'))
    self.setGeometry(100, 100, 900, 900)
    layout = QtWidgets.QVBoxLayout()
    layout.setSpacing(10)
        
    file_row = QHBoxLayout()
    file_row.setSpacing(5)
    self.refresh_button = QtWidgets.QPushButton('Refresh', self)
    self.refresh_button.setFixedWidth(90)
    self.refresh_button.clicked.connect(self.refresh_tables)
    self.file_input = QtWidgets.QLineEdit(self)
    self.file_input.setPlaceholderText('Enter Excel file path or use Browse')
    self.file_input.setMinimumWidth(400)
    self.browse_button = QtWidgets.QPushButton('Browse', self)
    self.browse_button.setFixedWidth(90)
    self.browse_button.clicked.connect(self.browse_file)
    file_row.addWidget(self.refresh_button)
    file_row.addWidget(self.file_input)
    file_row.addWidget(self.browse_button)
    layout.addLayout(file_row)
    
    date_row = QHBoxLayout()
    date_row.setSpacing(5)
    self.date_selector = QDateEdit(self)
    self.date_selector.setCalendarPopup(True)
    self.date_selector.setDate(QtCore.QDate.currentDate())
    self.date_selector.dateChanged.connect(self.update_selected_date)
    date_row.addWidget(QtWidgets.QLabel("Select Date:"))
    date_row.addWidget(self.date_selector)
    layout.addLayout(date_row)
    
    self.load_button = QtWidgets.QPushButton('Generate Pivot Table', self)
    self.load_button.clicked.connect(self.load_excel_file)
    layout.addWidget(self.load_button)
    self.pivot_table_display = QtWidgets.QTextEdit(self)
    self.pivot_table_display.setReadOnly(True)
    self.pivot_table_display.setMinimumHeight(500)
    layout.addWidget(self.pivot_table_display)
    
    self.status_bar = QtWidgets.QLabel("")
    layout.addWidget(self.status_bar)
    
    self.save_button = QtWidgets.QPushButton('Save', self)
    self.save_button.setToolTip("Save the generated tables to Excel")
    self.save_button.clicked.connect(self.save_excel_file)
    self.save_button.hide()
    layout.addWidget(self.save_button)
    
    self.send_email_button = QtWidgets.QPushButton('Send Pivot Table via Email', self)
    self.send_email_button.setToolTip("Send the generated tables as Excel via Outlook")
    self.send_email_button.clicked.connect(self.send_email)
    self.send_email_button.hide()
    layout.addWidget(self.send_email_button)
    
    self.setLayout(layout)
    
    self.file_watcher_timer = QtCore.QTimer()
    self.file_watcher_timer.setInterval(2000)
    self.file_watcher_timer.timeout.connect(self.check_file_changed)
    self.file_watcher_timer.start()
    
    def update_selected_date(self, date):
      self.selected_date = date.toPyDate()
    
    def browse_file(self):
      file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
      if file_path:
        self.file_input.setText(file_path)
        self.last_file_path = file_path
        self.last_file_mtime = os.path.getmtime(file_path)
        self.status_bar.setText("File selected. Click 'Generate Pivot Table' to view tables.")
        self.pivot_table_display.clear()
    
    def check_file_changed(self):
      file_path = self.file_input.text()
      if file_path and os.path.exists(file_path):
        mtime = os.path.getmtime(file_path)
        if self.last_file_path == file_path and self.last_file_mtime and mtime != self.last_file_mtime:
          self.last_file_mtime = mtime
          self.status_bar.setText("File changed on disk. Refreshing tables...")
          self.load_excel_file()
    
    def refresh_tables(self):
      self.status_bar.setText("Refreshing tables...")
      self.load_excel_file()
    
    def save_excel_file(self):
      file_path = self.file_input.text()
      try:
        df = pd.read_excel(file_path)
        df = map_column_names(df, [
          'Date', 'Client', 'Sow', 'No. of Employees', 'Start Date', 'End Date','Employee Code', '% Allocation', 'Allocation',
          'Project Code', 'Project Name','Project Type', 'Employee Name', 'Employee Job Title', 'Employee Grade', 'Location',
          'Home Department', 'Function', 'Function Hierarchy', 'Skill Item', 'Skill Item Category', 'Skill Item Category Group'
        ])
        df = add_skill_columns(df, file_path)
        if 'Client' not in df.columns:
          QMessageBox.critical(self, 'Error', "The 'Client' column was not found in the Excel file.")
          return
        
        company_names = df['Client'].dropna().unique()
        company_names = [name.strip().lower() for name in company_names if isinstance(name, str)]
      except Exception as e:
        QMessageBox.critical(self, 'Error', f"Error reading 'Client' column from Excel: {e}")
        return
      _, pivot_dfs, error = self.generate_and_display_tables(file_path)
      if error:
        QMessageBox.critical(self, 'Error', f"Error creating tables: {error}")
        return
      save_path, _ = QFileDialog.getSaveFileName(
        self,
        "Save Excel Attachment",
        f"Allocation_report_{datetime.today().date()}.xlsx",
        "Excel Files (*.xlsx)"
      )
      if not save_path:
        return
      try:
        with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
          guidelines_df = pd.DataFrame(GUIDELINES_COLUMNS)
          guidelines_df.to_excel(writer, sheet_name='Guidelines', index=False)
          
          worksheet = writer.sheets['Guidelines']
          for idx, col in enumerate(guidelines_df.columns):
            max_len = guidelines_df[col].astype(str).map(len).max()
            worksheet.set_column(idx, idx, max_len + 2)
            
            sheet_names = [
              "Client Allocation Summary",
              "Monthly Headcount Summary",
              "Total Headcount Summary",
              "Allocation Comparison",
              "Billed Employee Sheet",
              "Unbilled Employee Sheet",
              "Allocations Ending This Month"
            ]
            
            color_map = {client: PALETTE[i % len(PALETTE)] for i, client in enumerate(company_names)}
            header_format = writer.book.add_format({'bold': True, 'align': 'center', 'color': 'black'})                                                                    # pyright: ignore[reportAttributeAccessIssue]
            black_font_format = writer.book.add_format({'color': 'black', 'align': 'left'})                                                                                # pyright: ignore[reportAttributeAccessIssue]
            cell_format = writer.book.add_format({'align': 'center'})                                                                                                      # pyright: ignore[reportAttributeAccessIssue]
             
            for df, name in zip(pivot_dfs, sheet_names):
              sheet_df = df if df is not None and not df.empty else pd.DataFrame()
              if 'Skill Item' in df.columns and 'Skill Item Category' in df.columns:
                print(f"Saving {name} with skills columns included.")
              sheet_df.to_excel(writer, sheet_name=name, index=False)
              worksheet = writer.sheets[name]
              
              # Adjust column width for content visibility
              for col_idx, col in enumerate(sheet_df.columns):
                max_length = max(sheet_df[col].astype(str).map(len).max() if not sheet_df.empty else 0, len(col) + 5)
                worksheet.set_column(col_idx, col_idx, max_length, cell_format)
                if name not in ["Client Allocation Summary", "Monthly Headcount Summary", "Total Headcount Summary"]:
                  for client, color in color_map.items():
                    client_format = writer.book.add_format({'bg_color': color, 'align': 'center'})                                                                       # pyright: ignore[reportAttributeAccessIssue]
                    worksheet.conditional_format(
                      'A1:Z{}'.format(len(sheet_df) + 1),
                      {'type': 'formula', 'criteria': f'=$A1="{client}"', 'format': client_format}
                    )
                 if name == "Unbilled Employee Sheet":
                   for col_idx in range(len(sheet_df.columns)):
                     worksheet.write(0, col_idx, sheet_df.columns[col_idx], header_format)
                     if sheet_df.columns[col_idx] in ['Skill Item', 'Skill Item Category', 'Skill Item Category Group']:
                       worksheet.set_column(col_idx, col_idx, 25, black_font_format)
            QMessageBox.information(self, 'Success', 'Excel file saved successfully.')
        except Exception as e:
          QMessageBox.critical(self, 'Error', f"Error saving Excel file: {e}")
          return

    def generate_and_display_tables(self, file_path):
      try:
        df = pd.read_excel(file_path)
        df = map_column_names(df, [
          'Date', 'Client', 'Sow', 'No. of Employees', 'Start Date', 'End Date','Employee Code', '% Allocation', 'Allocation',
          'Project Code', 'Project Name','Project Type', 'Employee Name', 'Employee Job Title', 'Employee Grade', 'Location',
          'Home Department', 'Function', 'Function Hierarchy', 'Skill Item', 'Skill Item Category', 'Skill Item Category Group'
        ])
        df = add_skill_columns(df, file_path)
        if df is None:
          raise ValueError("DataFrame is not loaded.")
          
        df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce')
        current_date = pd.to_datetime("today")
        first_day_of_current_month = pd.Timestamp(current_date.year, current_date.month, 1)
        df = df[df['End Date'] >= first_day_of_current_month]
           
        history_file = os.path.join(os.path.dirname(file_path), "allocation_history.xlsx")
      
        # Table 1: Client Allocation Summary
        df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce')
        df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce')
        df['UniqueID'] = range(1, len(df) + 1)
        today = pd.to_datetime(self.selected_date)
        df['Today Date'] = today
        df = df[df['Start Date'] <= today]
            
        df['Allocation'] = df['Allocation'].where(df['Allocation'].isin(['Current Allocation', 'Future Allocation']))
            
        if not df.empty:
          if {'Client', 'Sow', 'Employee Code', 'UniqueID', 'Today Date', 'Allocation', 'Start Date', 'End Date'}.issubset(df.columns):
            df_alloc = df[['UniqueID', 'Client', 'Sow', 'Employee Code', 'Today Date', 'Allocation', 'Start Date', 'End Date']].dropna(
              subset=['Client', 'Sow', 'Employee Code', 'UniqueID']
            )
            df_alloc['Allocation Type'] = df_alloc['Sow'].apply(normalize_sow)
            df_alloc['Allocation Type'] = df_alloc['Allocation Type'].astype(SOW_CAT_TYPE)

            if '% Allocation' in df.columns:
              df_alloc = df_alloc.merge(df[['UniqueID', 'Employee Code', '% Allocation']], on='UniqueID', how='left')
              if 'Employee Code_x' in df_alloc.columns and 'Employee Code_y' in df_alloc.columns:
                df_alloc['Employee Code'] = df_alloc['Employee Code_x']
                df_alloc.drop(columns=['Employee Code_x', 'Employee Code_y'], inplace=True)
                df_alloc = df_alloc.drop_duplicates(subset=['Client', 'Employee Code', 'Allocation Type'])
                df_alloc['Adjusted Count'] = df_alloc['% Allocation'] / 100.0
                merged_table = df_alloc.groupby(['Client', 'Allocation Type'], observed=True)['Adjusted Count'].sum().reset_index()
                merged_table.columns = ['Client', 'Sow', 'No. of Employees']
              else:
                merged_table = df_alloc.groupby(['Client', 'Allocation Type'], observed=True)['Employee Code'].count().reset_index()
                merged_table.columns = ['Client', 'Sow', 'No. of Employees']
              merged_table['Sow'] = merged_table['Sow'].astype(SOW_CAT_TYPE)
              
              if APAC_CLIENT_NAME not in merged_table['Client'].values:
                apac_rows = pd.DataFrame({
                  'Client': [APAC_CLIENT_NAME] * len(PRIORITY_ORDER),
                  'Sow': PRIORITY_ORDER,
                  'No. of Employees': [0] * len(PRIORITY_ORDER)
                })
                apac_rows['Sow'] = apac_rows['Sow'].astype(SOW_CAT_TYPE)
                merged_table = pd.concat([merged_table, apac_rows], ignore_index=True)
                
              if (df_alloc['Today Date'] == today).all():
                update_allocation_history(merged_table, history_file, self.selected_date)
              else:
                QMessageBox.critical(self, 'Error', "Error generating tables: Outdated File Upload Detected")
                
              merged_table_pivot = merged_table.pivot_table(
                index='Client',
                columns='Sow',
                values='No. of Employees',
                aggfunc='sum',
                fill_value=0,
                observed=False
              )

              if APAC_CLIENT_NAME not in merged_table_pivot.index:
                merged_table_pivot.loc[APAC_CLIENT_NAME] = [0] * len(merged_table_pivot.columns)
                
              merged_table_pivot['Total Employees'] = merged_table_pivot.sum(axis=1).round(2)
              merged_table_pivot['Expected Utilization [(Pre-Sow + Billed) / Total]'] = ((
                merged_table_pivot['Billed'] + merged_table_pivot['PreSow']) / merged_table_pivot['Total Employees']
              ).round(2)
              merged_table_pivot['Billed% [Biiled / Total]'] = (
                merged_table_pivot['Billed'] / merged_table_pivot['Total Employees']
              ).round(2)
              else:
                merged_table_pivot = pd.DataFrame(columns=[*PRIORITY_ORDER, 'Total Employees'])
                return None, [None] * 6, "Missing required columns for Client allocation."
            else:
              merged_table_pivot = pd.DataFrame(columns=[*PRIORITY_ORDER, 'Total Employees'])
              df_alloc = pd.DataFrame(columns=['UniqueID', 'Client', 'Sow', 'Employee Code', 'Today Date', 'Allocation Type', 'Start Date', 'End Date'])
              merged_table = pd.DataFrame(columns=['Client', 'Sow', 'No. of Employees'])
              
            # Table 2: Monthly Headcount Summary
            if df_alloc is None or df_alloc.empty:
              headcount_table = pd.DataFrame(columns=['Client', 'Sow', 'Total Headcount'])
            else:
              first_day_of_month = pd.Timestamp(today.year, today.month, 1)
              last_day_of_month = first_day_of_month + pd.offsets.MonthEnd(0)
              df_alloc_monthly = df_alloc[
                # (df_alloc['End Date'] <= last_day_of_month) &
                # (df_alloc['End Date'].dt.month == first_day_of_month.month) &
                # (df_alloc['End Date'].dt.year == first_day_of_month.year)
                (df_alloc['Start Date'] <= today) &
                (df_alloc['End Date'] >= today)
              ]
              headcount_table = df_alloc_monthly.groupby(['Client', 'Allocation Type'], observed=True)['Employee Code'].nunique().reset_index()
              headcount_table.columns = ['Client', 'Sow', 'Total Headcount']
              
            ALL_CLIENTS = sorted(set(df_alloc['Client'].unique()) if df_alloc is not None and not df_alloc.empty else [])
            if APAC_CLIENT_NAME not in ALL_CLIENTS:
              ALL_CLIENTS.append(APAC_CLIENT_NAME)
            full_index = pd.MultiIndex.from_product([ALL_CLIENTS, PRIORITY_ORDER], names=['Client', 'Sow'])
            headcount_table = headcount_table.set_index(['Client', 'Sow']).reindex(full_index, fill_value=0).reset_index()
            headcount_table['Sow'] = headcount_table['Sow'].astype(SOW_CAT_TYPE)

            headcount_table_pivot = headcount_table.pivot_table(
              index='Client',
              columns='Sow',
              values='Total Headcount',
              aggfunc='sum',
              fill_value=0,
              observed=False
            )

            headcount_table_pivot['Total Headcount'] = headcount_table_pivot.sum(axis=1)
            total_row = pd.DataFrame(headcount_table_pivot.sum(axis=0)).T
            total_row.index = ['Total']                                                                                                                       # pyright: ignore[reportAttributeAccessIssue]
            headcount_table_pivot = pd.concat([headcount_table_pivot, total_row])
            headcount_table_pivot.index.name = 'Client'

            if 'Client' in merged_table.columns:
              existing_clients = set(merged_table['Client'].unique())
              missing_clients = set(ALL_CLIENTS) - existing_clients
              
              if missing_clients:
                missing_rows = pd.DataFrame({
                  'Client': list(missing_clients) * len(PRIORITY_ORDER),
                  'Sow': PRIORITY_ORDER * len(missing_clients),
                  'No. of Employees': [0] * len(PRIORITY_ORDER) * len(missing_clients)
                })
                missing_rows['Sow'] = missing_rows['Sow'].astype(SOW_CAT_TYPE)
                merged_table = pd.concat([merged_table, missing_rows], ignore_index=True)
                
            # Table 3: Total Headcount Summary
            if df_alloc is None or df_alloc.empty:
              headcount_table_today = pd.DataFrame(columns=['Client', 'Sow', 'Total Headcount'])
            else:
              df_alloc_filtered = df_alloc[(df_alloc['Start Date'] <= today) & (df_alloc['End Date'] >= today)]
              headcount_table_today = df_alloc_filtered.groupby(['Client', 'Allocation Type'], observed=True)['Employee Code'].nunique().reset_index()
              headcount_table_today.columns = ['Client', 'Sow', 'Total Headcount']
              
            headcount_table_today['Sow'] = headcount_table_today['Sow'].astype(SOW_CAT_TYPE)
            if APAC_CLIENT_NAME not in headcount_table_today['Client'].values:
              apac_headcount_rows = pd.DataFrame({
                'Client': [APAC_CLIENT_NAME] * len(PRIORITY_ORDER),
                'Sow': PRIORITY_ORDER,
                'Total Headcount': [0] * len(PRIORITY_ORDER)
              })
              apac_headcount_rows['Sow'] = apac_headcount_rows['Sow'].astype(SOW_CAT_TYPE)
              headcount_table_today = pd.concat([headcount_table_today, apac_headcount_rows], ignore_index=True)
              
            headcount_table_today_pivot = headcount_table_today.pivot_table(
              index='Client',
              columns='Sow',
              values='Total Headcount',
              aggfunc='sum',
              fill_value=0,
              observed=False
            )
            if APAC_CLIENT_NAME not in headcount_table_today_pivot.index:
              headcount_table_today_pivot.loc[APAC_CLIENT_NAME] = [0] * len(headcount_table_today_pivot.columns)
              
            headcount_table_today_pivot['Total Headcount'] = headcount_table_today_pivot.sum(axis=1)
            total_row = pd.DataFrame(headcount_table_today_pivot.sum(axis=0)).T
            total_row.index = ['Total']                                                                                     # pyright: ignore[reportAttributeAccessIssue]
            headcount_table_today_pivot = pd.concat([headcount_table_today_pivot, total_row])
            headcount_table_today_pivot.index.name = 'Client'

            if 'Client' in merged_table.columns:
              if APAC_CLIENT_NAME not in merged_table['Client'].values:
                apac_rows = pd.DataFrame({
                  'Client': [APAC_CLIENT_NAME] * len(PRIORITY_ORDER),
                  'Sow': PRIORITY_ORDER,
                  'No. of Employees': [0] * len(PRIORITY_ORDER)
                })
                apac_rows['Sow'] = apac_rows['Sow'].astype(SOW_CAT_TYPE)
                merged_table = pd.concat([merged_table, apac_rows], ignore_index=True)
                
            # Table 4: Allocation Comparison
            try:
              final_df = generate_employee_change_summary(history_file, merged_table)
            except Exception as exc:
              final_df = pd.DataFrame(columns=[
                'Client', 'Sow', f'Count for Today ({today.date()})',
                'Count for Previous day', 'Count for Previous week',
                'Change from Previous day', 'Change from Previous week'
              ])
              print(f"Warning: allocation comparison could not be generated: {exc}")
          
            # Tables 5, 6, 7: Employee Sheets
            required_cols_new = [
              'Client', 'Project Code', 'Project Name', 'Project Type',
              'Employee Code', 'Employee Name', 'Employee Job Title',
              'Employee Grade', 'Location', '% Allocation', 'Sow',
              'Start Date', 'End Date', 'Allocation', 'Home Department',
              'Function', 'Function Hierarchy','Skill Item',
              'Skill Item Category', 'Skill Item Category Group'
            ]

            if set(required_cols_new).issubset(df.columns):
              pivot_df_new = df[required_cols_new].dropna(how='all')
              pivot_df_new['Normalized Sow'] = pivot_df_new['Sow'].apply(normalize_sow)
              pivot_df_new['Start Date'] = pd.to_datetime(pivot_df_new['Start Date'], format='%d-%m-%Y', errors='coerce', dayfirst=True)
              pivot_df_new['End Date'] = pd.to_datetime(pivot_df_new['End Date'], format='%d-%m-%Y', errors='coerce', dayfirst=True)
            else:
              pivot_df_new = pd.DataFrame(columns=required_cols_new + ['Normalized Sow', 'Start Date', 'End Date'])
              
            today_ts = pd.Timestamp.today()
            current_year = today_ts.year
            current_month = today_ts.month
            month_days = calendar.monthrange(current_year, current_month)[1]

            if not pivot_df_new.empty:
              pivot_df_new['Adjusted Start Date'] = pivot_df_new['Start Date'].apply(
                lambda x: x if (pd.notna(x) and x.month == current_month and x.year == current_year) else pd.Timestamp(current_year, current_month, 1)
              )
              pivot_df_new['Adjusted End Date'] = pivot_df_new['End Date'].apply(
                lambda x: x if (pd.notna(x) and x.month == current_month and x.year == current_year) else pd.Timestamp(current_year, current_month, month_days)
              )
              pivot_df_new['#Day'] = (pivot_df_new['Adjusted End Date'] - pivot_df_new['Adjusted Start Date']).dt.days
              pivot_df_new['#MonthDays'] = month_days
              if '% Allocation' not in pivot_df_new.columns:
                pivot_df_new['% Allocation'] = 0
                pivot_df_new['#CalendarDay'] = ((pivot_df_new['#Day'] / pivot_df_new['#MonthDays']) * pivot_df_new['% Allocation']).round(2)
                pivot_df_new['Aging'] = (pivot_df_new['Start Date'] - today_ts).dt.days                                                                 # pyright: ignore[reportOperatorIssue]
              else:
                pivot_df_new['#Day'] = []
                pivot_df_new['#MonthDays'] = []
                pivot_df_new['#CalendarDay'] = []
                pivot_df_new['Aging'] = []
                
            # Table 5: Billed Employee Sheet
            billed_cols = [
              'Client', 'Normalized Sow', 'Project Code', 'Project Name',
              'Employee Code', 'Employee Name', 'Employee Job Title',
              'Employee Grade', 'Location', '% Allocation', 'Sow',
              'Start Date', 'End Date', '#Day', '#CalendarDay', 'Allocation'
            ]
            pivot_df_billed = pivot_df_new[pivot_df_new['Normalized Sow'] == 'Billed'][billed_cols].reset_index(drop=True)
            if not pivot_df_billed.empty:
              pivot_df_billed['Start Date'] = pd.to_datetime(pivot_df_billed['Start Date'], errors='coerce', dayfirst=True).dt.strftime('%d-%m-%Y')
              pivot_df_billed['End Date'] = pd.to_datetime(pivot_df_billed['End Date'], errors='coerce', dayfirst=True).dt.strftime('%d-%m-%Y')

            # Table 6: Unbilled Employee Sheet
            non_billed_cols = [
              'Client', 'Normalized Sow', 'Project Code', 'Project Name',
              'Employee Code', 'Employee Name', 'Employee Job Title',
              'Employee Grade', 'Location', '% Allocation', 'Sow',
              'Start Date', 'End Date', 'Home Department', 'Function',
              'Function Hierarchy', '#Day', '#CalendarDay', 'Aging', 'Allocation',
              'Skill Item', 'Skill Item Category', 'Skill Item Category Group'
            ]
            pivot_df_non_billed = pivot_df_new[pivot_df_new['Normalized Sow'] != 'Billed'][non_billed_cols].reset_index(drop=True)
            if not pivot_df_non_billed.empty:
              pivot_df_non_billed['Aging'] = (pivot_df_non_billed['Start Date'] - today_ts).dt.days                                                     # pyright: ignore[reportOperatorIssue]
              pivot_df_non_billed['Start Date'] = pd.to_datetime(pivot_df_non_billed['Start Date'], errors='coerce', dayfirst=True).dt.strftime('%d-%m-%Y')
              pivot_df_non_billed['End Date'] = pd.to_datetime(pivot_df_non_billed['End Date'], errors='coerce', dayfirst=True).dt.strftime('%d-%m-%Y')
            
            # Table 7: Employees whose allocation ends in the current month
            required_cols_5 = [
              'Client', 'Project Code', 'Project Name', 'Project Type',
              'Employee Code', 'Employee Name', 'Employee Job Title',
              'Employee Grade', 'Sow', 'Start Date', 'End Date', 'Allocation'
            ]

            if set(required_cols_5).issubset(df.columns):
              df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce', dayfirst=True)
              df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce', dayfirst=True)
              df['Normalized Sow'] = df['Sow'].apply(normalize_sow)
              df = df[df['End Date'].notna()]
              
              mask = (
              (df['End Date'].dt.month == today_ts.month) &
              (df['End Date'].dt.year == today_ts.year) &
              (df['Normalized Sow'] == 'Billed') &
              (df['Allocation'] != 'Future allocation') &
              (df['Project Type'] != 'Non Billable')
            )
            
            pivot_df5 = df.loc[mask, required_cols_5].copy()
            pivot_df5['Start Date'] = pivot_df5['Start Date'].dt.strftime('%d-%m-%Y')
            pivot_df5['End Date'] = pivot_df5['End Date'].dt.strftime('%d-%m-%Y')
          else:
            pivot_df5 = pd.DataFrame(columns=required_cols_5)
            
          merged_table_pivot = merged_table_pivot.reset_index()
          headcount_table_pivot = headcount_table_pivot.reset_index()
          headcount_table_today_pivot = headcount_table_today_pivot.reset_index()
          final_df = final_df.sort_values(by=['Client', 'Sow'])
          pivot_df_billed = pivot_df_billed.sort_values(by=['Client', 'Sow']).reset_index(drop=True) if not pivot_df_billed.empty else pivot_df_billed
          pivot_df_non_billed = pivot_df_non_billed.sort_values(by=['Client', 'Sow']).reset_index(drop=True) if not pivot_df_non_billed.empty else pivot_df_non_billed
          pivot_df5 = pivot_df5.sort_values(by=['Client', 'Sow']).reset_index(drop=True) if not pivot_df5.empty else pivot_df5

          html_merged = df_to_html_with_row_color(merged_table_pivot, 'Client', apply_color=False)
          html_headcount = df_to_html_with_row_color(headcount_table_pivot, 'Client', apply_color=False)
          html_headcount_today = df_to_html_with_row_color(headcount_table_today_pivot, 'Client', apply_color=False)
          html_compare = df_to_html_with_row_color(final_df, 'Client')
          html_billed = df_to_html_with_row_color(pivot_df_billed, 'Client')
          html_non_billed = df_to_html_with_row_color(pivot_df_non_billed, 'Client')
          html5 = df_to_html_with_row_color(pivot_df5, 'Client')
          combined_html = (
            "<h3>Table 1: Client Allocation Summary</h3>" + html_merged +
            "<br><h3>Table 2: Monthly Headcount Summary</h3>" + html_headcount +
            "<br><h3>Table 3: Total Headcount Summary</h3>" + html_headcount_today +
            "<br><h3>Table 4: Allocation Comparison</h3>" + html_compare +
            "<br><h3>Table 5: Billed Employee Details</h3>" + html_billed +
            "<br><h3>Table 6: Unbilled Employee Details</h3>" + html_non_billed +
            "<br><h3>Table 7: Employees whose project end in current month</h3>" + html5
          )

          return (
            combined_html,
            [
              merged_table_pivot,
              headcount_table_pivot,
              headcount_table_today_pivot,
              final_df,
              pivot_df_billed[billed_cols] if not pivot_df_billed.empty else pd.DataFrame(columns=billed_cols),
              pivot_df_non_billed[non_billed_cols] if not pivot_df_non_billed.empty else pd.DataFrame(columns=non_billed_cols),
              pivot_df5[required_cols_5] if not pivot_df5.empty else pd.DataFrame(columns=required_cols_5)
            ],
            None
          )
        except Exception as e:
          return None, [None] * 7, str(e)

    def load_excel_file(self):
      file_path = self.file_input.text()
      if not file_path:
        QMessageBox.warning(self, 'Input Error', 'Please enter file path.')
        self.status_bar.setText("No file selected.")
        return
        
      self.pivot_table_display.setHtml("<i>Loading, please wait...</i>")
      self.status_bar.setText("Generating tables, please wait...")
      QtWidgets.QApplication.processEvents()
      
      worker = TableWorker(file_path, self.generate_and_display_tables)
      worker.signals.finished.connect(self.on_tables_ready)
      self.threadpool.start(worker)

    def on_tables_ready(self, combined_html, pivot_dfs, error):
      if error:
        QMessageBox.critical(self, 'Error', error)
        self.pivot_table_display.setPlainText(error)
        self.status_bar.setText("Error: " + error)
        self.pivot_df = None
        return
        
      self.pivot_table_display.setHtml(combined_html)
      self.pivot_df = pd.concat([df for df in pivot_dfs if df is not None and not df.empty], ignore_index=True) if any(df is not None and not df.empty for df in pivot_dfs) else pd.DataFrame()
      self.status_bar.setText("Tables generated successfully.")
      self.save_button.show()
      self.send_email_button.show()
      print("All tables generated and printed.")
      
    def send_email(self):
      subject = "Latest Allocation Report"
      file_path = self.file_input.text()
      try:
        df = pd.read_excel(file_path)
        df = map_column_names(df, [
          'Date', 'Client', 'Sow', 'No. of Employees', 'Start Date', 'End Date','Employee Code', '% Allocation', 'Allocation',
          'Project Code', 'Project Name','Project Type', 'Employee Name', 'Employee Job Title', 'Employee Grade', 'Location',
          'Home Department', 'Function', 'Function Hierarchy', 'Skill Item', 'Skill Item Category', 'Skill Item Category Group'
        ])
        if df is None:
          raise ValueError("DataFrame is not loaded.")
          
        all_company_names = df['Client'].dropna().unique()
        all_company_names = [normalize_company_name(name) for name in all_company_names if isinstance(name, str)]
        
        recipient_to_companies = {}
        for company, recipients in COMPANY_RECIPIENTS.items():
          norm_name = normalize_company_name(company)
          for recipient in recipients:
            recipient_to_companies.setdefault(recipient, set()).add(norm_name)
            
        universal_recipient_companies = set(normalize_company_name(company) for company in COMPANY_RECIPIENTS.keys())
        for universal_recipient in UNIVERSAL_RECIPIENTS:
          recipient_to_companies[universal_recipient] = universal_recipient_companies
          
        _, pivot_dfs_full, error = self.generate_and_display_tables(file_path)
        if error:
          QMessageBox.critical(self, 'Error', f"Error creating tables: {error}")
          return
          
        global_unbilled = pivot_dfs_full[5] if pivot_dfs_full and len(pivot_dfs_full) >= 6 else pd.DataFrame()
        current_date_str = datetime.today().date().strftime('%d-%m-%Y')
        guidelines_df = pd.DataFrame(GUIDELINES_COLUMNS)
        sheet_names = [
          "Client Allocation Summary",
          "Monthly Headcount Summary",
          "Total Headcount Summary",
          "Allocation Comparison",
          "Billed Employee Sheet",
          "Unbilled Employee Sheet",
          "Allocations Ending This Month"
        ]
        
        for recipient, companies in recipient_to_companies.items():
          if not any(df_item['Client'].apply(lambda x: normalize_company_name(x)).isin(companies).any() if df_item is not None else False for df_item in pivot_dfs_full):
            print(f"Skipping email for {recipient} as no data available.")
            continue
            
          if recipient in UNIVERSAL_RECIPIENTS:
            company_name_display = ', '.join(sorted(COMPANY_RECIPIENTS.keys()))
            pivot_dfs = pivot_dfs_full
          else:                    
            company_name_display = ', '.join(                        
              sorted(company for company in COMPANY_RECIPIENTS if normalize_company_name(company) in companies)                    
            )
            pivot_dfs = [
              df_item[df_item['Client'].apply(lambda x: normalize_company_name(x)).isin(companies)]
              if df_item is not None else pd.DataFrame()
              for idx, df_item in enumerate(pivot_dfs_full)
            ]
            pivot_dfs[5] = global_unbilled
            
            headcount_df_filtered_month = pivot_dfs[1]                
            headcount_summary_month_html = "<p><b>Monthly Headcount Summary:</b></p><table border='1' cellpadding='4' style='border-collapse:collapse; font-size:12px; text-align:center;'>"
            headcount_summary_month_html += "<tr>" + "".join(f"<th>{col}</th>" for col in headcount_df_filtered_month.columns) + "</tr>"
            for _, row in headcount_df_filtered_month.iterrows():
              row_html = "".join(f"<td>{row[col]}</td>" for col in headcount_df_filtered_month.columns)
              headcount_summary_month_html += f"<tr>{row_html}</tr>"
            headcount_summary_month_html += "</table>"
            
            headcount_df_filtered_today = pivot_dfs[2]
            headcount_summary_today_html = "<p><b>Total Headcount Summary:</b></p><table border='1' cellpadding='4' style='border-collapse:collapse; font-size:12px; text-align:center;'>"
            headcount_summary_today_html += "<tr>" + "".join(f"<th>{col}</th>" for col in headcount_df_filtered_today.columns) + "</tr>"
            for _, row in headcount_df_filtered_today.iterrows():
              row_html = "".join(f"<td>{row[col]}</td>" for col in headcount_df_filtered_today.columns)
              headcount_summary_today_html += f"<tr>{row_html}</tr>"
            headcount_summary_today_html += "</table>"
            
            temp_file_path = os.path.join(tempfile.gettempdir(), f"allocation_report_{recipient.replace('@','_').replace('.','_')}_{current_date_str}.xlsx")
            color_map = {client: PALETTE[i % len(PALETTE)] for i, client in enumerate(company_name_display.split(', '))}
            
            with pd.ExcelWriter(temp_file_path, engine='xlsxwriter') as writer:
              guidelines_df.to_excel(writer, sheet_name='Guidelines', index=False)
              worksheet = writer.sheets['Guidelines']
              for idx, col in enumerate(guidelines_df.columns):
                max_len = guidelines_df[col].astype(str).map(len).max()
                worksheet.set_column(idx, idx, max_len + 2)
              cell_format = writer.book.add_format({'align': 'center'})                                                            # pyright: ignore[reportAttributeAccessIssue]
              for df_item, name in zip(pivot_dfs, sheet_names):
                sheet_df = df_item if df_item is not None and not df_item.empty else pd.DataFrame()
                sheet_df.to_excel(writer, sheet_name=name, index=False)
                worksheet = writer.sheets[name]
                for col_idx, col in enumerate(sheet_df.columns):
                  max_length = max(sheet_df[col].astype(str).map(len).max() if not sheet_df.empty else 0, len(col) + 5)
                  worksheet.set_column(col_idx, col_idx, max_length, cell_format)
                if name not in ["Client Allocation Summary", "Monthly Headcount Summary", "Total Headcount Summary"]:
                  for client, color in color_map.items():
                    client_format = writer.book.add_format({'bg_color': color, 'align': 'center'})                                 # pyright: ignore[reportAttributeAccessIssue]
                    worksheet.conditional_format(
                      'A1:Z{}'.format(len(sheet_df) + 1),
                      {'type': 'formula', 'criteria': f'=$A1="{client}"', 'format': client_format}
                    )
            body = (
              f"<div style='font-size:12px;'>"
              f"Hi,<br><br>"
              f"Please find attached the latest allocation report for <b>{company_name_display or 'your company'}</b>.<br><br>"
              f"The report includes the following sheets:<br>"
              f"&emsp;&emsp;1. Client Allocation Summary<br>"
              f"&emsp;&emsp;2. Monthly Headcount Summary<br>"
              f"&emsp;&emsp;3. Total Headcount Summary<br>"
              f"&emsp;&emsp;4. Allocation Comparison<br>"
              f"&emsp;&emsp;5. Billed Employee Database<br>"
              f"&emsp;&emsp;6. Unbilled Employee Database<br>"
              f"&emsp;&emsp;7. Employees whose projects end in this month<br><br>"
              f"{headcount_summary_month_html}<br><br>"
              f"{headcount_summary_today_html}<br><br>"
              f"If you have any questions or need further analysis, please feel free to reach out.<br><br>"
              f"Best regards,<br>"
              f"{'Sender name'}<br>"
              f"</div>"
            )
              
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.Subject = subject
            mail.To = recipient
            mail.Cc = "; ".join(UNIVERSAL_RECIPIENTS) if recipient not in UNIVERSAL_RECIPIENTS else ""
            mail.Recipients.ResolveAll()
            mail.HTMLBody = body
            mail.Attachments.Add(temp_file_path)
            mail.Send()
            
            # Clean up temp files
            with suppress(Exception):
              os.remove(temp_file_path)
              
            QMessageBox.information(self, 'Success', 'Emails sent successfully.')
      except Exception as e:
        QMessageBox.critical(self, 'Error', str(e))
# -------------------------
# Run Application
# -------------------------
if __name__ == '__main__':
  app = QtWidgets.QApplication(sys.argv)
  window = ExcelPivotApp()
  window.show()
  sys.exit(app.exec_())
