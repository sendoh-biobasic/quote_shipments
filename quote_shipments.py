import pandas as pd
import win32com.client as win32
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import datetime


class FreightApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Freight Quotation Automation")
        self.root.geometry("600x550")

        tk.Label(root, text="Step 1: Select Template File", font=('Arial', 10, 'bold')).pack(pady=(10, 0))
        self.file_frame = tk.Frame(root)
        self.file_frame.pack(pady=5)
        self.file_path_entry = tk.Entry(self.file_frame, width=50)
        self.file_path_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(self.file_frame, text="Browse", command=self.browse_file).pack(side=tk.LEFT)

        tk.Label(root, text="Step 2: Enter Date (YYYY-MM-DD)", font=('Arial', 10, 'bold')).pack(pady=(10, 0))
        self.date_entry = tk.Entry(root, width=25)
        self.date_entry.insert(0, datetime.today().strftime('%Y-%m-%d'))
        self.date_entry.pack(pady=5)

        self.btn_run = tk.Button(root, text="START PROCESS", bg="#4CAF50", fg="white",
                                 font=('Arial', 11, 'bold'), width=20, height=2, command=self.run_logic)
        self.btn_run.pack(pady=15)

        self.log_area = scrolledtext.ScrolledText(root, width=70, height=15, font=('Consolas', 9))
        self.log_area.pack(pady=5, padx=10)

    def log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.root.update_idletasks()

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if filename:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, filename)

    def run_logic(self):
        file_path = self.file_path_entry.get()
        target_date_str = self.date_entry.get().strip()

        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid Excel file!")
            return

        self.log_area.delete(1.0, tk.END)
        self.log(f"INFO: Opening {os.path.basename(file_path)}...")

        try:
            # 1. Read sheets (no header, use integer column indices)
            df_raw = pd.read_excel(file_path, sheet_name='Sheet1', header=None)
            df_db  = pd.read_excel(file_path, sheet_name='database', header=None)

            # 2. Parse target date
            try:
                target_date_val = pd.to_datetime(target_date_str).strftime('%Y-%m-%d')
            except Exception:
                self.log(f"ERROR: Invalid input date: {target_date_str}")
                return

            # 3. Find rows in Sheet1 matching the target date (col A = index 0)
            valid_rows = []
            for i, row in df_raw.iterrows():
                cell_date = pd.to_datetime(row[0], errors='coerce')
                if pd.notna(cell_date) and cell_date.strftime('%Y-%m-%d') == target_date_val:
                    valid_rows.append(row)

            if not valid_rows:
                self.log(f"WARNING: No records found for {target_date_val}")
                self.log("DEBUG: First 10 rows of column A:")
                for i in range(min(10, len(df_raw))):
                    self.log(f"  Row {i + 1}: {df_raw.iloc[i, 0]}")
                return

            # 4. Build lookup maps from database sheet
            # database columns (0-indexed):
            #   0=A Location key, 1=B Address, 2=C Primary courier, 3=D Secondary courier
            #   6=G Courier name (for email lookup), 7=H Email address
            df_db_clean = df_db.iloc[1:].copy()           # skip blank row 0
            df_db_clean = df_db_clean.dropna(subset=[0])  # skip rows with no location key
            df_db_clean[0] = df_db_clean[0].astype(str).str.strip()

            # addr_map: location key -> full address string (col B)
            addr_map = pd.Series(
                df_db_clean[1].values,
                index=df_db_clean[0]
            ).to_dict()

            # location_couriers_map: location key -> [courier names]
            # Reads col C (index 2) as primary courier, col D (index 3) as secondary courier
            # Both are included if they have values
            location_couriers_map = {}
            for _, r in df_db_clean.iterrows():
                key = str(r[0]).strip()
                couriers = []
                for col_idx in [2, 3]:  # col C then col D
                    val = str(r[col_idx]).strip() if pd.notna(r[col_idx]) else ""
                    if val and val != "nan":
                        couriers.append(val)
                location_couriers_map[key] = couriers

            # courier_email_map: courier name (col G) -> email (col H)
            courier_email_map = {}
            for _, r in df_db_clean.iterrows():
                g_val = str(r[6]).strip() if pd.notna(r[6]) else ""
                h_val = str(r[7]).strip() if pd.notna(r[7]) else ""
                if g_val and g_val != "nan" and h_val and h_val != "nan":
                    courier_email_map[g_val] = h_val

            # 5. Send emails via Outlook
            outlook = win32.Dispatch('outlook.application')
            success_count = 0

            for row in valid_rows:
                # Sheet1 column mapping (0-indexed):
                # 0=Date, 1=From, 2=To, 8=Details, 9=Memo, 10=# of SKIDS
                from_key = str(row[1]).strip()
                to_key   = str(row[2]).strip()

                # Collect all couriers from col C + D for the matched location (From or To)
                # Each courier gets its own email; duplicates are skipped
                couriers_to_send = []
                for loc_key in [from_key, to_key]:
                    if loc_key in location_couriers_map:
                        for courier_name in location_couriers_map[loc_key]:
                            email = courier_email_map.get(courier_name, "")
                            if email and (courier_name, email) not in couriers_to_send:
                                couriers_to_send.append((courier_name, email))

                if not couriers_to_send:
                    self.log(f"SKIP: No courier info for '{from_key}' or '{to_key}'")
                    continue

                from_addr    = addr_map.get(from_key, from_key)
                to_addr      = addr_map.get(to_key,   to_key)
                num_skids    = row[10] if pd.notna(row[10]) else ""
                memo         = row[9]  if pd.notna(row[9])  else ""
                details      = row[8]  if pd.notna(row[8])  else ""
                # Preserve line breaks from Excel cell into HTML
                details_html = str(details).replace('\n', '<br>')

                html_body = f"""
                <html><body style="font-family: Calibri;">
                    <p>Hello,</p>
                    <p>Please send quotation on freight for the following shipment to services@biobasic.com.</p>
                    <table>
                        <tr><td><b>FROM/PICK UP:</b></td><td>{from_addr}</td></tr>
                        <tr><td><b>TO/DELIVERY ADDRESS:</b></td><td>{to_addr}</td></tr>
                        <tr><td><b>DETAILS:</b></td><td>{num_skids} Skids</td></tr>
                        <tr><td><b>MEMO:</b></td><td>{memo}</td></tr>
                        <tr><td><b>DETAILS:</b></td><td>{details_html}</td></tr>
                    </table>
                </body></html>
                """

                # Send one email per courier (col C courier + col D courier if present)
                for courier_name, recipient_email in couriers_to_send:
                    mail = outlook.CreateItem(0)
                    mail.To = recipient_email
                    mail.Subject = "Quotation"
                    mail.HTMLBody = html_body
                    mail.Send()
                    success_count += 1
                    self.log(f"SUCCESS: Sent to {courier_name} <{recipient_email}>")

            messagebox.showinfo("Done", f"Sent {success_count} emails.")

        except Exception as e:
            self.log(f"CRITICAL ERROR: {str(e)}")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = FreightApp(root)
    root.mainloop()