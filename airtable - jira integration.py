#!/usr/bin/env python
# combined_diff_airtable.py
import re
import argparse
import csv
import json
import requests
import pandas as pd
import sys
import os
import datetime

def map_field(field):
    """If field starts with 'Custom field', return the value inside parentheses; otherwise, return field unchanged."""
    m = re.match(r"Custom field\s*\((.+)\)", field)
    if m:
        return m.group(1).strip()
    return field

def chunks(lst, n):
    # Yield successive n-sized chunks from lst
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

def fetch_airtable_data(url, headers):
    # Retrieve all records from Airtable (handling pagination) and return a list of records.
    all_records = []
    params = {}
    while True:
        response = requests.get(url, headers=headers, params=params)
        if response.status_code != 200:
            print("Error fetching Airtable data:", response.text)
            sys.exit(1)
        data = response.json()
        all_records.extend(data.get("records", []))
        if "offset" in data:
            params["offset"] = data["offset"]
        else:
            break
    return all_records

def backup_to_excel(records, backup_path):
    """
    Write the fetched Airtable records (each record's "fields" dict) to an Excel backup file.
    Uses fixed column widths for columns A through E:
      A = 13, B = 11, C = 15, D = 30, E = 60, F = 16.
    Any additional columns are set to a default width of 20.
    """
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

def main():
    parser = argparse.ArgumentParser(
        description="Pull current data from Airtable, compare it to NEW CSV file, generate diff log (Excel and JSON), and update Airtable via API."
    )
    parser.add_argument("--new", required=True, help="Path to the new CSV file")
    parser.add_argument("--excel", required=True, help="Path for the output Excel workbook")
    parser.add_argument("--json", required=True, help="Path for the output JSON file (changes payload)")
    parser.add_argument("--backup", default="airtable_backup.xlsx", help="Base name for the Airtable backup Excel file")
    parser.add_argument("--base_id", required=True, help="Airtable base ID")
    parser.add_argument("--table", required=True, help="Airtable table name or ID")
    parser.add_argument("--token", required=True, help="Airtable API token")
    parser.add_argument("--preview", "-p", action="store_true",
                        help="Perform all operations except the Airtable API upload (for customer preview)")
    parser.add_argument("--upload", action="store_true",
                        help="Skip diff generation; use an existing JSON payload for upload only")
    args = parser.parse_args()

    base_id = args.base_id
    table_name = args.table
    api_token = args.token
    url = f"https://api.airtable.com/v0/{base_id}/{table_name}"
    headers = {
        "Authorization": f"Bearer {api_token}",
        "Content-Type": "application/json"
    }

    # Today's date for the API update field.
    today_date = datetime.datetime.now().strftime("%Y-%m-%d")

    if not args.upload:
        if not (args.new and args.excel):
            parser.error("Full mode requires --new and --excel arguments. Use --upload to skip diff generation.")
        airtable_records = fetch_airtable_data(url, headers)

        # Convert fetched records into a dictionary keyed by "Issue key"
        old_data = {}
        for rec in airtable_records:
            fields = rec.get("fields", {})
            issue_key = fields.get("Issue key")
            if issue_key:
                old_data[issue_key] = fields

        changes = []            # For Change Log
        changed_rows_count = 0  # Count of loggable changes to existing tickets
        new_tickets_count = 0   # Count of new tickets (not in old_data)
        records = []            # For JSON upsert payload
        new_ticket_records = [] # For New Tickets worksheet

        # Include "Created" among fields to check.
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
            "Created"  # NEW: Include "Created"
        ]
        field_mapping = {
            "Custom field (T-Shirt Size)": "T-shirt Size",
            "Custom field (Customer Reported)": "Customer Reported"
        }

        # Read CSV and combine duplicate "Labels" columns.
        csv_rows = []
        with open(args.new, newline='', encoding='utf-8') as new_file:
            reader = csv.reader(new_file)
            headers_list = next(reader)
            if "Issue key" not in headers_list:
                print("Issue key column not found in new CSV header!")
                sys.exit(1)
            for row in reader:
                row_dict = {}
                for i, header in enumerate(headers_list):
                    key = header.strip()
                    value = row[i].strip() if i < len(row) else ""
                    if key:
                        if key in row_dict:
                            if key == "Labels" and value:
                                row_dict[key] = row_dict[key] + "," + value
                        else:
                            row_dict[key] = value
                csv_rows.append(row_dict)

        # Compute set of Jira issue keys from CSV.
        jira_issue_keys = { row.get("Issue key", "") for row in csv_rows if row.get("Issue key", "") }
        # Identify tickets in Airtable not present in the new Jira export.
        smart_tickets_not_in_jira = [ old_data[k] for k in old_data if k not in jira_issue_keys ]
        smart_not_count = len(smart_tickets_not_in_jira)

        # Process each CSV row.
        for row_dict in csv_rows:
            labels = row_dict.get("Labels", "")
            new_quarter = ""
            if "#2026" in labels:
                new_quarter = "Q1_2026"
            elif "#1_p_sm" in labels:
                new_quarter = "Q1"
            elif "#2_p_sm" in labels:
                new_quarter = "Q2"
            elif "#3_p_sm" in labels:
                new_quarter = "Q3"
            elif "#4_p_sm" in labels:
                new_quarter = "Q4"
            if new_quarter and new_quarter != "Q1_2026":
                new_quarter = f"{new_quarter}_{datetime.datetime.now().strftime('%Y')}"
            new_matrix = "Yes" if "24rvw_matrix_parity" in labels else "No"

            issue_key = row_dict.get("Issue key", "")
            if not issue_key:
                continue

            record_fields = {"Issue key": issue_key}
            record_changed = False   # Any change for JSON payload
            loggable_change = False  # Change that should trigger an API update (and changelog)

            if issue_key in old_data:
                old_row = old_data[issue_key]
                # Process each field in one loop.
                for field in fields_to_check:
                    mapped_field = field_mapping.get(field, field)
                    csv_val = row_dict.get(mapped_field)
                    if csv_val is None:
                        csv_val = row_dict.get(field, "")
                    if mapped_field == "Fix versions" and csv_val.startswith("smartmls-connectmls-"):
                        csv_val = re.sub(r"^smartmls-connectmls-", "", csv_val)
                    # For "Created", convert CSV value to datetime and compare with Airtable value.
                    if mapped_field == "Created" and csv_val:
                        try:
                            dt_csv = datetime.datetime.strptime(csv_val, "%m/%d/%Y %H:%M")
                        except ValueError:
                            dt_csv = None
                        try:
                            old_created = old_row.get("Created", "")
                            try:
                                dt_air = datetime.datetime.strptime(old_created, "%Y-%m-%dT%H:%M:%S.%fZ")
                            except ValueError:
                                dt_air = datetime.datetime.strptime(old_created, "%Y-%m-%dT%H:%M:%SZ")
                        except ValueError:
                            dt_air = None
                        if dt_csv and dt_air and dt_csv == dt_air:
                            csv_val = old_created  # Force CSV value to match Airtable
                    old_val = old_row.get(mapped_field, "")
                    if csv_val != old_val:
                        record_changed = True
                        record_fields[mapped_field] = csv_val
                        if mapped_field not in ["Created", "Labels", "Last Updated (API)"]:
                            loggable_change = True
                            changes.append({
                                "Ticket Number": issue_key,
                                "Column": mapped_field,
                                "prev": old_val,
                                "new": csv_val
                            })
                # Process Quarter (always loggable)
                old_quarter = old_row.get("Quarter", "")
                if new_quarter != old_quarter:
                    record_fields["Quarter"] = new_quarter
                    record_changed = True
                    loggable_change = True
                    changes.append({
                        "Ticket Number": issue_key,
                        "Column": "Quarter",
                        "prev": old_quarter,
                        "new": new_quarter
                    })
                # Process Matrix Parity (always loggable)
                old_matrix = old_row.get("Matrix Parity", "")
                if new_matrix != old_matrix:
                    record_fields["Matrix Parity"] = new_matrix
                    record_changed = True
                    loggable_change = True
                    changes.append({
                        "Ticket Number": issue_key,
                        "Column": "Matrix Parity",
                        "prev": old_matrix,
                        "new": new_matrix
                    })
                record_fields.pop("Labels", None)
                # Only add this record if there is a loggable change.
                if loggable_change:
                    record_fields["Last Updated (API)"] = today_date
                    records.append({"fields": record_fields})
                    changed_rows_count += 1
            else:
                # New ticket branch: always include new tickets.
                new_tickets_count += 1
                record_fields["Quarter"] = new_quarter
                record_fields["Matrix Parity"] = new_matrix
                for field in fields_to_check:
                    mapped_field = field_mapping.get(field, field)
                    csv_val = row_dict.get(mapped_field)
                    if csv_val is None:
                        csv_val = row_dict.get(field, "")
                    if mapped_field == "Fix versions" and csv_val.startswith("smartmls-connectmls-"):
                        csv_val = re.sub(r"^smartmls-connectmls-", "", csv_val)
                    # For new tickets, leave "Created" as provided.
                    if mapped_field == "Labels":
                        record_fields[mapped_field] = csv_val
                        continue
                    record_fields[mapped_field] = csv_val
                record_fields.pop("Labels", None)
                record_fields["Last Updated (API)"] = today_date
                records.append({"fields": record_fields})
                new_ticket_records.append(record_fields)

        if not records:
            print("There are no differences between the current Airtable data and the data it is being compared to. Exiting the operation.")
            sys.exit(0)
        else:
            print(f"There were {changed_rows_count} changed rows between this Jira export and the current Airtable")
            print(f"Net new tickets since last export: {new_tickets_count}")
            print(f"Smart Tickets Not in Jira: {smart_not_count}")

        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base_backup, _ = os.path.splitext(args.backup)
        backup_file = f"{base_backup}_{now_str}.xlsx"
        backup_to_excel(airtable_records, backup_file)

        # Create DataFrames for Excel output.
        overview_data = [
            {"": "Tickets with changes since last export", "Count": changed_rows_count},
            {"": "Net new tickets since last export", "Count": new_tickets_count},
            {"": "Smart Tickets Not in Jira", "Count": smart_not_count}
        ]
        df_overview = pd.DataFrame(overview_data)
        df_changes = pd.DataFrame(changes)

        with pd.ExcelWriter(args.excel, engine='xlsxwriter') as writer:
            df_overview.to_excel(writer, sheet_name="Overview", index=False)
            df_changes.to_excel(writer, sheet_name="Change Log", index=False)
            if smart_not_count > 0:
                df_smart = pd.DataFrame(smart_tickets_not_in_jira)
                df_smart.to_excel(writer, sheet_name="Smart Tickets Not in Jira", index=False)
            if new_tickets_count > 0:
                df_new = pd.DataFrame(new_ticket_records)
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
            if not df_changes.empty and len(df_changes.columns) > 0:
                changes_ws.write(0, 0, df_changes.columns[0], header_format)
                changes_ws.write(0, 1, df_changes.columns[1], header_format)

            col_widths = [11, 11, 13, 21, 60, 12, 16, 11, 16, 18, 12, 18, 60]
            smart_sheet = writer.sheets.get("Smart Tickets Not in Jira")
            if smart_sheet is not None:
                for i, width in enumerate(col_widths):
                    smart_sheet.set_column(i, i, width)
            new_sheet = writer.sheets.get("New Tickets")
            if new_sheet is not None:
                for i, width in enumerate(col_widths):
                    new_sheet.set_column(i, i, width)

        print("- Excel workbook written to", args.excel)

        payload = {
            "performUpsert": {
                "fieldsToMergeOn": ["Issue key"]
            },
            "records": records
        }

        with open(args.json, "w", encoding="utf-8") as json_file:
            json.dump(payload, json_file, indent=4)
        print("- JSON payload written to", args.json)

    else:
        print("Upload-only mode enabled. Using existing JSON payload from:", args.json)
        with open(args.json, "r", encoding="utf-8") as json_file:
            payload = json.load(json_file)

    if args.preview:
        print("- Preview enabled. Changelog generated but data not pushed to Airtable.\n- Rerun program with the '--upload' flag to push to Airtable.")
        return

    confirm = input(f"WARNING: Airtable upload triggered. This will modify table '{table_name}' at Base Id '{base_id}'. Type 'yes' or 'y' to proceed: ")
    if confirm.lower() not in ['yes', 'y']:
        print("Upload canceled.")
        sys.exit(0)

    records_list = payload.get("records", [])
    chunk_size = 10

    for idx, chunk in enumerate(chunks(records_list, chunk_size), start=1):
        chunk_payload = {
            "performUpsert": {
                "fieldsToMergeOn": ["Issue key"]
            },
            "records": chunk,
            "typecast": True  # Enable typecast in case of select fields.
        }
        print(f"Sending chunk {idx} with {len(chunk)} records...")
        response = requests.patch(url, headers=headers, json=chunk_payload)
        print("Status Code:", response.status_code)
        try:
            print("Response:", response.json())
        except json.JSONDecodeError:
            print("Response could not be decoded as JSON:", response.text)

if __name__ == '__main__':
    main()
