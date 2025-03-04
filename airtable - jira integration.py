#!/usr/bin/env python
import re
import csv
import json
import requests
import pandas as pd
import sys
import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# -----------------------------
# Constants for Airtable credentials
# -----------------------------
BASE_ID = "<>"
TABLE_ID = "<>"
API_TOKEN = "<>"

# -----------------------------
# Helper Functions
# -----------------------------
def map_field(field):
    m = re.match(r"Custom field\s*\((.+)\)", field)
    if m:
        return m.group(1).strip()
    return field

def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

def fetch_airtable_data(url, headers):
    all_records = []
    params = {}
    while True:
        response = requests.get(url, headers=headers, params=params)
        if response.status_code != 200:
            messagebox.showerror("Error", f"Error fetching Airtable data:\n{response.text}")
            return None
        data = response.json()
        all_records.extend(data.get("records", []))
        if "offset" in data:
            params["offset"] = data["offset"]
        else:
            break
    return all_records

def detect_duplicate_keys(airtable_records):
    freq = {}
    duplicates = set()
    for rec in airtable_records:
        fields = rec.get("fields", {})
        key = fields.get("Issue key")
        if key:
            freq[key] = freq.get(key, 0) + 1
    for key, count in freq.items():
        if count > 1:
            duplicates.add(key)
    return duplicates

def backup_to_excel(records, backup_path):
    data = [rec.get("fields", {}) for rec in records]
    df = pd.DataFrame(data)
    with pd.ExcelWriter(backup_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Backup", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Backup"]
        fixed_widths = [13, 11, 15, 30, 60, 16]
        default_width = 20
        for idx, col in enumerate(df.columns):
            width = fixed_widths[idx] if idx < len(fixed_widths) else default_width
            worksheet.set_column(idx, idx, width)
    print("- Airtable backup Excel written to:", backup_path)

def read_and_normalize_csv(csv_path):
    rows = []
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        headers = next(reader)
        for row in reader:
            row_dict = {}
            for i, header in enumerate(headers):
                key = header.strip()
                value = row[i].strip() if i < len(row) else ""
                if key:
                    # Combine duplicate 'Labels' columns
                    if key in row_dict and key == "Labels" and value:
                        row_dict[key] = row_dict[key] + "," + value
                    else:
                        row_dict.setdefault(key, value)
            rows.append(row_dict)
    return rows

def build_old_data(airtable_records):
    out = {}
    for rec in airtable_records:
        fields = rec.get("fields", {})
        issue_key = fields.get("Issue key")
        if issue_key:
            out[issue_key] = fields
    return out

def normalize_fix_versions(value):
    if value.startswith("smartmls-connectmls-"):
        return re.sub(r"^smartmls-connectmls-", "", value)
    return value

def convert_created(csv_val):
    try:
        dt_csv = datetime.datetime.strptime(csv_val, "%d/%b/%y %I:%M %p")
    except ValueError:
        return csv_val
    return dt_csv.strftime("%m-%d-%Y %H:%M")

def unify_bool(csv_val, old_val):
    val_lower = csv_val.strip().lower()
    if not val_lower:
        return None, None
    if val_lower == "yes":
        new_val = True
    elif val_lower == "no":
        new_val = False
    else:
        return None, None
    if isinstance(old_val, bool):
        old_bool = old_val
    elif isinstance(old_val, str):
        low = old_val.strip().lower()
        old_bool = True if low in ["yes", "true"] else False
    else:
        old_bool = False
    return new_val, old_bool

def unify_old_bool(old_val):
    if isinstance(old_val, bool):
        return old_val
    if isinstance(old_val, str):
        low = old_val.strip().lower()
        if low in ["yes","true"]:
            return True
    return False

def unify_matrix_parity(labels, old_val):
    new_val = "24rvw_matrix_parity" in labels
    if isinstance(old_val, bool):
        old_bool = old_val
    elif isinstance(old_val, str):
        low = old_val.strip().lower()
        old_bool = True if low in ["yes", "true"] else False
    else:
        old_bool = False
    return new_val, old_bool

def compare_record(csv_row, old_row, fields_to_check, field_mapping, last_updated, changelog):
    record_fields = {}
    record_changed = False
    loggable_change = False

    for field in fields_to_check:
        mapped = field_mapping.get(field, field)
        if mapped == "Created":
            continue
        csv_val = csv_row.get(mapped, csv_row.get(field, ""))
        old_val = old_row.get(mapped, "")
        
        if mapped == "Fix versions":
            csv_val = normalize_fix_versions(csv_val)
        if mapped == "Resolution":
            csv_val = csv_val.strip()
            if not csv_val:
                continue
        elif mapped == "Customer Reported":
            new_bool, old_bool = unify_bool(csv_val, old_val)
            if new_bool is None:
                continue
            if new_bool != old_bool:
                record_changed = True
                record_fields[mapped] = new_bool
                loggable_change = True
                changelog.append({
                    "Ticket Number": csv_row.get("Issue key", ""),
                    "Column": mapped,
                    "prev": old_bool,
                    "new": new_bool
                })
            continue

        if csv_val != old_val:
            record_changed = True
            record_fields[mapped] = csv_val
            if mapped not in ["Labels", "Last Updated (API)"]:
                loggable_change = True
                changelog.append({
                    "Ticket Number": csv_row.get("Issue key", ""),
                    "Column": mapped,
                    "prev": old_val,
                    "new": csv_val
                })

    labels = csv_row.get("Labels", "")
    new_quarter = ""
    if "#2026" in labels:
        new_quarter = "Q1_UTC"
    elif "#1_p_sm" in labels:
        new_quarter = "Q1"
    elif "#2_p_sm" in labels:
        new_quarter = "Q2"
    elif "#3_p_sm" in labels:
        new_quarter = "Q3"
    elif "#4_p_sm" in labels:
        new_quarter = "Q4"
    if new_quarter and new_quarter != "Q1_UTC":
        new_quarter = f"{new_quarter}_{datetime.datetime.now().strftime('%Y')}"
    old_quarter = old_row.get("Quarter", "")
    if new_quarter != old_quarter:
        record_fields["Quarter"] = new_quarter
        record_changed = True
        loggable_change = True
        changelog.append({
            "Ticket Number": csv_row.get("Issue key", ""),
            "Column": "Quarter",
            "prev": old_quarter,
            "new": new_quarter
        })

    new_matrix, old_matrix = unify_matrix_parity(labels, old_row.get("Matrix Parity", ""))
    if new_matrix != old_matrix:
        record_fields["Matrix Parity"] = new_matrix
        record_changed = True
        loggable_change = True
        changelog.append({
            "Ticket Number": csv_row.get("Issue key", ""),
            "Column": "Matrix Parity",
            "prev": old_matrix,
            "new": new_matrix
        })

    record_fields["Issue key"] = csv_row.get("Issue key", "")
    record_fields.pop("Labels", None)
    if record_changed and loggable_change:
        record_fields["Last Updated (API)"] = last_updated

    return record_changed, loggable_change, record_fields

def compare_new_record(csv_row, fields_to_check, field_mapping, last_updated):
    record_fields = {}
    labels = csv_row.get("Labels", "")

    for field in fields_to_check:
        mapped = field_mapping.get(field, field)
        csv_val = csv_row.get(mapped, csv_row.get(field, ""))

        if mapped == "Fix versions":
            csv_val = normalize_fix_versions(csv_val)

        if mapped == "Created" and csv_val:
            csv_val = convert_created(csv_val)

        if mapped == "Resolution":
            csv_val = csv_val.strip()
            if not csv_val:
                continue

        if mapped == "Customer Reported":
            new_bool, _ = unify_bool(csv_val, False)
            if new_bool is None:
                continue
            record_fields[mapped] = new_bool
        else:
            record_fields[mapped] = csv_val

    new_quarter = ""
    if "#2026" in labels:
        new_quarter = "Q1_UTC"
    elif "#1_p_sm" in labels:
        new_quarter = "Q1"
    elif "#2_p_sm" in labels:
        new_quarter = "Q2"
    elif "#3_p_sm" in labels:
        new_quarter = "Q3"
    elif "#4_p_sm" in labels:
        new_quarter = "Q4"
    if new_quarter and new_quarter != "Q1_UTC":
        new_quarter = f"{new_quarter}_{datetime.datetime.now().strftime('%Y')}"
    record_fields["Quarter"] = new_quarter

    new_matrix, _ = unify_matrix_parity(labels, False)
    record_fields["Matrix Parity"] = new_matrix

    record_fields["Issue key"] = csv_row.get("Issue key", "")
    record_fields.pop("Labels", None)
    record_fields["Last Updated (API)"] = last_updated
    return record_fields

def output_excel(excel_path, overview_data, df_changelog, new_data, col_widths):
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        pd.DataFrame(overview_data).to_excel(writer, sheet_name="Overview", index=False)
        df_changelog.to_excel(writer, sheet_name="Change Log", index=False)

        if new_data:
            df_new = pd.DataFrame([nr["fields"] for nr in new_data])
            df_new.to_excel(writer, sheet_name="New Tickets", index=False)

        workbook = writer.book
        overview_ws = writer.sheets["Overview"]
        changes_ws = writer.sheets["Change Log"]

        overview_ws.set_column('A:A', 35)
        overview_ws.set_column('B:B', 15)
        changes_ws.set_column('A:A', 16)
        changes_ws.set_column('B:B', 11)
        changes_ws.set_column('C:C', 50)
        changes_ws.set_column('D:D', 50)

        header_format = workbook.add_format({'align': 'left', 'bold': True})
        if not df_changelog.empty and len(df_changelog.columns) > 0:
            changes_ws.write(0, 0, df_changelog.columns[0], header_format)
            changes_ws.write(0, 1, df_changelog.columns[1], header_format)

        band_formats = [
            workbook.add_format({'bg_color': '#FFFFFF'}),
            workbook.add_format({'bg_color': '#E6E6E6'})
        ]

        prev_ticket = None
        band_idx = 0
        for row_idx in range(len(df_changelog)):
            current_ticket = df_changelog.iloc[row_idx]["Ticket Number"]
            if current_ticket != prev_ticket:
                band_idx = 1 - band_idx
            changes_ws.set_row(row_idx + 1, None, band_formats[band_idx])
            prev_ticket = current_ticket

        sheet_new = writer.sheets.get("New Tickets")
        if sheet_new is not None:
            for i, width in enumerate(col_widths):
                sheet_new.set_column(i, i, width)
    print("- Excel workbook written to", excel_path)

# -----------------------------
# GUI Application
# -----------------------------
class DiffApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Airtable Diff & Upload")
        # Change from 600x450 to something wider, e.g. 700x450:
        self.geometry("710x450")  
        # Optionally, also set a minimum size to prevent shrinking below a certain width/height:
        self.minsize(710, 450)
        self.create_widgets()

    def create_widgets(self):
        pad_options = {'padx': 5, 'pady': 5}
        frame = ttk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True)

        # Jira Export File (CSV)
        ttk.Label(frame, text="Jira Export File (CSV):").grid(row=0, column=0, sticky=tk.W, **pad_options)
        self.jira_csv_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.jira_csv_var, width=50).grid(row=0, column=1, **pad_options)
        ttk.Button(frame, text="Browse", command=self.browse_jira_csv).grid(row=0, column=2, **pad_options)

        # Show the Airtable credentials as labels, not editable
        ttk.Label(frame, text=f"Airtable Base ID: {BASE_ID}").grid(row=1, column=0, columnspan=2, sticky=tk.W, **pad_options)
        ttk.Label(frame, text=f"Airtable Table ID: {TABLE_ID}").grid(row=2, column=0, columnspan=2, sticky=tk.W, **pad_options)
        ttk.Label(frame, text=f"Airtable API Token: {API_TOKEN}").grid(row=3, column=0, columnspan=2, sticky=tk.W, **pad_options)

        # Checkboxes for preview and upload-only modes
        self.preview_var = tk.BooleanVar()
        ttk.Checkbutton(frame, text="Preview mode (don’t push to Airtable)", variable=self.preview_var).grid(row=4, column=0, columnspan=2, sticky=tk.W, **pad_options)

        self.upload_only_var = tk.BooleanVar()
        ttk.Checkbutton(frame, text="Upload-only mode (skip diff generation)", variable=self.upload_only_var).grid(row=5, column=0, columnspan=2, sticky=tk.W, **pad_options)

        # Run button
        ttk.Button(frame, text="Run", command=self.run_process).grid(row=6, column=0, columnspan=3, pady=15)

        # Log output text box
        ttk.Label(frame, text="Log Output:").grid(row=7, column=0, sticky=tk.W, **pad_options)
        self.log_text = tk.Text(frame, height=10)
        self.log_text.grid(row=8, column=0, columnspan=3, sticky="nsew", **pad_options)
        frame.rowconfigure(8, weight=1)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        print(message)

    def browse_jira_csv(self):
        file_path = filedialog.askopenfilename(
            title="Select Jira Export CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.jira_csv_var.set(file_path)

    def run_process(self):
        # Gather inputs from UI
        jira_csv = self.jira_csv_var.get().strip()
        preview = self.preview_var.get()
        upload_only = self.upload_only_var.get()

        # Hard-coded from constants
        base_id = BASE_ID
        table_id = TABLE_ID
        api_token = API_TOKEN

        # Validate required inputs
        if not jira_csv:
            messagebox.showerror("Input Error", "Please select the Jira Export file (CSV).")
            return

        # Auto-generate output filenames based on today's date
        today = datetime.date.today()
        excel_out = f"smartMLS_changelog_{today.month}-{today.day}.xlsx"
        json_out = f"smartMLS_payload_{today.month}-{today.day}.json"

        # Get cwd for output messages
        working_dir = os.getcwd()
        excel_full_path = os.path.join(working_dir, excel_out)
        json_full_path = os.path.join(working_dir, json_out)

        url = f"https://api.airtable.com/v0/{base_id}/{table_id}"
        headers = {
            "Authorization": f"Bearer {api_token}",
            "Content-Type": "application/json"
        }
        last_updated = datetime.datetime.now(datetime.timezone.utc).strftime("%m-%d-%Y %H:%M")
        json_payload = None

        try:
            if not upload_only:
                self.log("Fetching Airtable data...")
                airtable_records = fetch_airtable_data(url, headers)
                if airtable_records is None:
                    return

                duplicate_keys = detect_duplicate_keys(airtable_records)
                if duplicate_keys:
                    dup_str = "\n".join(duplicate_keys)
                    self.log("WARNING: The following duplicate issue keys were found in Airtable and will be skipped:\n" + dup_str)

                old_data = build_old_data(airtable_records)
                self.log("Reading Jira Export CSV file...")
                csv_rows = read_and_normalize_csv(jira_csv)

                fields_to_check = [
                    "Issue Type",
                    "Summary",
                    "Status",
                    "Custom field (T-Shirt Size)",
                    "Priority",
                    "Components",
                    "Custom field (Customer Reported)",
                    "Fix versions",
                    "Labels",
                    "Created",
                    "Resolution"
                ]
                field_mapping = {
                    "Custom field (T-Shirt Size)": "T-shirt Size",
                    "Custom field (Customer Reported)": "Customer Reported"
                }

                changed_records = []
                new_records = []
                changelog_entries = []
                closed_statuses = {"In Prod", "Closed"}

                for row in csv_rows:
                    issue_key = row.get("Issue key", "")
                    if not issue_key or issue_key in duplicate_keys:
                        continue
                    if issue_key in old_data:
                        rec_changed, loggable, rec_fields = compare_record(
                            row, old_data[issue_key],
                            fields_to_check, field_mapping,
                            last_updated, changelog_entries
                        )
                        if rec_changed and loggable:
                            changed_records.append({"fields": rec_fields})
                    else:
                        status_val = row.get("Status", "").strip()
                        if status_val in closed_statuses:
                            continue
                        new_ticket_fields = compare_new_record(
                            row, fields_to_check, field_mapping, last_updated
                        )
                        new_records.append({"fields": new_ticket_fields})

                total_changed = len(changed_records)
                total_new = len(new_records)
                self.log(f"Changed rows: {total_changed}")
                self.log(f"New tickets: {total_new}")

                if total_changed == 0 and total_new == 0:
                    messagebox.showinfo("No Changes", "There are no differences. Exiting operation.")
                    return

                # Create backup file (auto-generated)
                now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_file = f"{os.path.splitext('airtable_backup.xlsx')[0]}_{now_str}.xlsx"
                backup_file_full_path = os.path.join(working_dir, backup_file)
                self.log("Creating Airtable backup...")
                backup_to_excel(airtable_records, backup_file)
                self.log(f"Backup file created at: {backup_file_full_path}")

                # Generate Excel changelog
                overview_data = [
                    {"": "Tickets with changes since last export", "Count": total_changed},
                    {"": "Net new tickets since last export", "Count": total_new}
                ]
                df_changelog = pd.DataFrame(changelog_entries)
                col_widths = [11, 11, 13, 21, 60, 12, 16, 11, 16, 18, 12, 18, 60]
                self.log("Generating Excel change log...")
                output_excel(excel_out, overview_data, df_changelog, new_records, col_widths)
                self.log(f"Change log generated at: {excel_full_path} ")

                # Write JSON payload
                json_payload = {
                    "performUpsert": {"fieldsToMergeOn": ["Issue key"]},
                    "records": changed_records + new_records
                }
                with open(json_out, "w", encoding="utf-8") as jf:
                    json.dump(json_payload, jf, indent=4)
                self.log("JSON payload generated at: " + json_full_path)
            else:
                # Upload-only mode: use existing JSON file
                self.log("Upload-only mode enabled. Loading existing JSON payload...")
                if not os.path.exists(json_out):
                    messagebox.showerror("Missing File", f"JSON file '{json_out}' not found for upload-only mode.")
                    return
                with open(json_out, "r", encoding="utf-8") as jf:
                    json_payload = json.load(jf)

            if preview:
                self.log("Preview mode enabled. No data will be pushed to Airtable.")
                messagebox.showinfo("Preview Mode", "Changelog generated but data not pushed to Airtable.\nRerun without preview mode to push changes.")
                return

            # Final confirmation
            warn_msg = (
                f"WARNING: Airtable upload triggered.\n"
                f"This will modify table '{table_id}' at Base ID '{base_id}'.\n"
                f"Do you want to proceed?"
            )
            if not messagebox.askyesno("Confirm Upload", warn_msg):
                self.log("Upload canceled by user.")
                return

            records_list = json_payload.get("records", [])
            chunk_size = 10
            for idx, chunk in enumerate(chunks(records_list, chunk_size), start=1):
                chunk_payload = {
                    "performUpsert": {"fieldsToMergeOn": ["Issue key"]},
                    "records": chunk,
                    "typecast": True
                }
                self.log(f"Uploading chunk {idx} with {len(chunk)} records...")
                response = requests.patch(url, headers=headers, json=chunk_payload)
                self.log("Status Code: " + str(response.status_code))
                try:
                    self.log("Response: " + str(response.json()))
                except json.JSONDecodeError:
                    self.log("Response could not be decoded as JSON: " + response.text)
            self.log("Upload complete.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.log("Error: " + str(e))

if __name__ == '__main__':
    app = DiffApp()
    app.mainloop()
