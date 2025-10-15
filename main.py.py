#!/usr/bin/env python3
"""
BLE Device Scanner with Batch Processing (cleaned single-file refactor + full debug logging)

This file preserves your application's behavior and adds detailed logging across
the codebase. The log file is overwritten on each start and saved as
BLE_6v2_debug_log.txt in the script directory.
"""

import threading
import time
import os
import csv
import pyperclip
import serial
import serial.tools.list_ports
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import win32print
from datetime import datetime
from queue import Queue

# Bartender COM automation (Windows).
import win32com.client
import pythoncom

import sys
# import traceback

# --- Logging setup ---------------------------------------------------------
DEBUG_LOG = "debug_log.txt"

def log_debug(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(DEBUG_LOG, "a") as f:
        f.write(f"[{timestamp}] {message}\n")

# --- Constants / Defaults ---------------------------------------------------
IGNORED_DEVICES = set()
SCANNED_DEVICES = {}         # { mac: { 'name', 'rssi', 'serial_no' } }
csv_serial_counter = 1       # incrementing SL.No for CSV rows
PRINT_DELAY = 1.0            # seconds default
# ---------------------------------------------------------------------------


class BartenderPrinter:
    """Persistent Bartender instance that prints labels one by one."""

    def __init__(self, btw_path):
        log_debug(f"[BT] Initializing Bartender for path: {btw_path}")
        self.bt = None
        self.format = None
        self.btw_path = btw_path
        
        try:
            pythoncom.CoInitialize()
            log_debug("[BT] Attempting to create Bartender.Application COM object")
            self.bt = win32com.client.Dispatch("Bartender.Application")
            log_debug(f"[BT] Bartender.Application created: {self.bt}")
            
            # Try both Visible settings to see what works
            self.bt.Visible = False  # Set to True for debugging
            log_debug("[BT] Bartender Visible property set to True")
            
            if btw_path and os.path.exists(btw_path):
                try:
                    log_debug(f"[BT] Opening format: {btw_path}")
                    self.format = self.bt.Formats.Open(btw_path, False, "")
                    log_debug(f"[BT] Format opened successfully: {self.format}")
                    
                    # Test if we can access the named substrings
                    try:
                        named_substrings = self.format.NamedSubStrings
                        log_debug(f"[BT] Named substrings available: {named_substrings.Count}")
                        for i in range(1, named_substrings.Count + 1):
                            log_debug(f"[BT] Substring {i}: {named_substrings.Item(i).Name}")
                    except Exception as e:
                        log_debug(f"[BT] Could not enumerate named substrings: {e}")
                        
                except Exception as e:
                    log_debug(f"[BT] Error opening BTW file: {e}")
                    messagebox.showerror("Bartender Error", f"Cannot open label file:\n{e}")
            else:
                log_debug(f"[BT] BTW path invalid or not provided: {btw_path}")
                
            log_debug("[BT] Bartender initialization complete")
        except Exception as e:
            log_debug(f"[BT] Error initializing Bartender COM: {e}")
            messagebox.showerror("Bartender Error", f"Cannot initialize Bartender:\n{e}")

    def print_label(self, mac, pcb_id):
        log_debug(f"[BT] print_label called with MAC={mac}, PCB={pcb_id}")
        
        if not self.format:
            log_debug("[BT] No format available; cannot print")
            messagebox.showwarning("Print Error", "No label format loaded")
            return False
            
        try:
            log_debug("[BT] Setting named substrings...")
            # Try setting the values
            self.format.SetNamedSubStringValue("MAC", mac)
            self.format.SetNamedSubStringValue("ID", pcb_id)
            log_debug("[BT] Named substrings set successfully")
            
            log_debug("[BT] Calling PrintOut...")
            self.format.PrintOut(False, False)  # ShowDialog=False, PrintToFile=False
            log_debug("[BT] PrintOut called successfully")
            
            # Check if there are any print errors
            try:
                print_jobs = self.bt.PrintJobs
                log_debug(f"[BT] Active print jobs: {print_jobs.Count}")
            except Exception as e:
                log_debug(f"[BT] Could not check print jobs: {e}")
                
            return True
            
        except Exception as e:
            log_debug(f"[BT] Print error: {e}")
            messagebox.showerror("Print Error", f"Failed to print label:\n{e}")
            return False

    def close(self):
        log_debug("[BT] close called")
        try:
            if self.format:
                try:
                    self.format.Close(1)  # btDoNotSaveChanges
                    log_debug("[BT] Format closed")
                except Exception as e:
                    log_debug(f"[BT] Error closing format: {e}")
            if self.bt:
                try:
                    self.bt.Quit()
                    log_debug("[BT] Bartender quit")
                except Exception as e:
                    log_debug(f"[BT] Error quitting Bartender: {e}")
        except Exception as e:
            log_debug(f"[BT] Unexpected error during close: {e}")

class App:
    def __init__(self, root):
        log_debug("App.__init__ start")
        # App state
        self.root = root
        self.btw_file_path = ""
        self.csv_file_path = ""
        self.version = ""
        self.code_version = ""
        self.csv_entries = []
        self.last_selected_port = ""
        self.ser_connection = None
        self.serial_thread = None
        self.is_scanning = False
        self.current_serial_number = 1

        # Persisted form data
        self.saved_version = ""
        self.saved_code_version = ""
        self.saved_btw_path = ""
        self.saved_csv_path = ""
        self.saved_batch_size = 0

        # Batch variables
        self.batch_size = 0
        self.batch_entries = []  # each entry: dict with pcb_id, mac, device_name, row_index, ui refs

        # Print queue and Bartender wrapper
        self.print_queue = Queue()
        self._bt_printer = None
        threading.Thread(target=self._print_worker, daemon=True).start()
        log_debug("Print worker thread started from __init__")

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Widgets container
        self.widgets = type("W", (), {})()

        # Build UI
        self._build_main_ui()
        self.show_config_page()
        log_debug("App.__init__ end")

    # -----------------------
    # UI construction
    # -----------------------
    def _build_main_ui(self):
        log_debug("_build_main_ui called")
        self.root.title("BLE Device Scanner with Batch Processing")
        self.root.geometry("820x800")

        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.pack(fill="both", expand=True)

        self.status_label = ttk.Label(
            self.root, text="Please configure the application to start.", font=("Arial", 9)
        )
        self.status_label.pack(side="bottom", fill="x", pady=(2, 0))

    # -----------------------
    # Page: Configuration
    # -----------------------
    def show_config_page(self):
        log_debug("show_config_page called")
        try:
            for w in self.main_frame.winfo_children():
                w.destroy()

            ttk.Label(self.main_frame, text="Configuration", font=("Arial", 16, "bold")).pack(pady=(0, 20))

            # Bartender selection
            btw_frame = ttk.Frame(self.main_frame)
            btw_frame.pack(fill="x", pady=5)
            ttk.Label(btw_frame, text="Bartender File:", width=15).pack(side="left")
            btw_label = ttk.Label(btw_frame, text="No file selected", foreground="gray")
            btw_label.pack(side="left", padx=(5, 10))

            def choose_btw():
                log_debug("choose_btw called (file dialog)")
                path = filedialog.askopenfilename(
                    title="Select Bartender Label File",
                    filetypes=[("Bartender Files", "*.btw"), ("All Files", "*.*")]
                )
                if path:
                    log_debug(f"Selected BTW file: {path}")
                    self.btw_file_path = path
                    btw_label.config(text=os.path.basename(path))
                else:
                    log_debug("choose_btw: no file chosen")

            ttk.Button(btw_frame, text="Select BTW File", command=choose_btw).pack(side="left")

            # Version / Code Version
            version_frame = ttk.Frame(self.main_frame)
            version_frame.pack(fill="x", pady=5)
            ttk.Label(version_frame, text="Version:", width=15).pack(side="left")
            self.widgets.version_entry = ttk.Entry(version_frame, width=30)
            self.widgets.version_entry.pack(side="left", padx=(5, 0))

            code_frame = ttk.Frame(self.main_frame)
            code_frame.pack(fill="x", pady=5)
            ttk.Label(code_frame, text="Code Version:", width=15).pack(side="left")
            self.widgets.code_entry = ttk.Entry(code_frame, width=30)
            self.widgets.code_entry.pack(side="left", padx=(5, 0))

            # CSV select or create
            csv_frame = ttk.Frame(self.main_frame)
            csv_frame.pack(fill="x", pady=5)
            ttk.Label(csv_frame, text="CSV File:", width=15).pack(side="left")
            csv_label = ttk.Label(csv_frame, text="No file selected", foreground="gray")
            csv_label.pack(side="left", padx=(5, 10))

            # Restore saved fields if present
            if self.saved_code_version:
                log_debug(f"Restoring saved_code_version: {self.saved_code_version}")
                self.widgets.code_entry.insert(0, self.saved_code_version)
            if self.saved_version:
                log_debug(f"Restoring saved_version: {self.saved_version}")
                self.widgets.version_entry.insert(0, self.saved_version)
            if self.saved_btw_path:
                log_debug(f"Restoring saved_btw_path: {self.saved_btw_path}")
                self.btw_file_path = self.saved_btw_path
                btw_label.config(text=os.path.basename(self.saved_btw_path))
            if self.saved_csv_path:
                log_debug(f"Restoring saved_csv_path: {self.saved_csv_path}")
                self.csv_file_path = self.saved_csv_path
                csv_label.config(text=os.path.basename(self.saved_csv_path))

            def select_csv():
                log_debug("select_csv called")
                path = filedialog.askopenfilename(
                    title="Select CSV File",
                    filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
                )
                if not path:
                    log_debug("select_csv: no path selected")
                    return

                expected_header = ["SL.No", "VERSION", "CODE VERSION", "ID", "MAC ID", "COMMENTS"]
                try:
                    with open(path, "r", newline="") as f:
                        reader = csv.reader(f)
                        header = next(reader, None)
                        if header is None:
                            log_debug("select_csv: Selected file is empty.")
                            messagebox.showerror("Error", "Selected file is empty. Please create a new CSV.")
                            return

                        header_clean = [h.strip() for h in header]
                        if header_clean != expected_header:
                            log_debug(f"select_csv: Invalid header found: {header_clean}")
                            messagebox.showerror(
                                "Invalid File",
                                "The selected CSV does not match the expected format.\n\n"
                                f"Expected: {expected_header}\nFound: {header_clean}\n\n"
                                "Please select the correct file or create a new one."
                            )
                            return

                    self.csv_file_path = path
                    csv_label.config(text=os.path.basename(path))
                    log_debug(f"CSV selected: {path}")
                    self.load_existing_csv_entries()
                    self.set_status(f"Loaded existing CSV: {os.path.basename(path)}")

                except Exception as e:
                    log_debug(f"select_csv: Failed to read CSV file: {e}")
                    messagebox.showerror("Error", f"Failed to read CSV file: {e}")

            def create_csv():
                log_debug("create_csv called")
                path = filedialog.asksaveasfilename(
                    title="Create New CSV File",
                    filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
                    defaultextension=".csv"
                )
                if path:
                    log_debug(f"Creating new CSV at: {path}")
                    self.csv_file_path = path
                    csv_label.config(text=os.path.basename(path))
                    # create file with header
                    with open(path, "w", newline="") as f:
                        writer = csv.writer(f)
                        writer.writerow(["SL.No", "VERSION", "CODE VERSION", "ID", "MAC ID", "COMMENTS"])
                    self.load_existing_csv_entries()

            ttk.Button(csv_frame, text="Select CSV", command=select_csv).pack(side="left", padx=(0, 5))
            ttk.Button(csv_frame, text="Create New CSV", command=create_csv).pack(side="left")

            # Save and continue
            def save_and_continue():
                log_debug("save_and_continue called")
                self.version = self.widgets.version_entry.get().strip()
                self.code_version = self.widgets.code_entry.get().strip()

                if not self.version or not self.code_version:
                    log_debug("save_and_continue: Version or Code Version missing")
                    messagebox.showerror("Error", "Please enter both Version and Code Version")
                    return
                if not self.btw_file_path:
                    log_debug("save_and_continue: BTW file not selected")
                    messagebox.showerror("Error", "Please select a Bartender label file")
                    return
                if not self.csv_file_path:
                    log_debug("save_and_continue: CSV not selected")
                    messagebox.showerror("Error", "Please select or create a CSV file")
                    return

                self.saved_version = self.version
                self.saved_code_version = self.code_version
                self.saved_btw_path = self.btw_file_path
                self.saved_csv_path = self.csv_file_path
                log_debug(f"Saved config: version={self.saved_version}, code={self.saved_code_version}, btw={self.saved_btw_path}, csv={self.saved_csv_path}")

                self.show_batch_setup_page()

            ttk.Button(self.main_frame, text="Save and Continue", command=save_and_continue).pack(pady=20)
        except Exception as e:
            log_debug(f"Exception in show_config_page: {e}")

    # -----------------------
    # Page: Batch setup
    # -----------------------
    def show_batch_setup_page(self):
        log_debug("show_batch_setup_page called")
        try:
            for w in self.main_frame.winfo_children():
                w.destroy()

            ttk.Label(self.main_frame, text="Batch Setup", font=("Arial", 16, "bold")).pack(pady=(0, 20))

            frame = ttk.Frame(self.main_frame)
            frame.pack(fill="x", pady=20)

            ttk.Label(frame, text="Number of boards to process:", font=("Arial", 12)).pack(pady=(0, 10))
            self.widgets.batch_entry = ttk.Entry(frame, width=10, font=("Arial", 12), justify="center")
            self.widgets.batch_entry.pack(pady=5)

            if self.saved_batch_size:
                log_debug(f"Restoring saved_batch_size: {self.saved_batch_size}")
                self.widgets.batch_entry.insert(0, str(self.saved_batch_size))

            def start_batch():
                log_debug("start_batch called")
                try:
                    n = int(self.widgets.batch_entry.get().strip())
                    if n <= 0:
                        log_debug("start_batch: non-positive number provided")
                        messagebox.showerror("Error", "Please enter a positive number")
                        return
                    self.batch_size = n
                    self.batch_entries = []
                    for i in range(self.batch_size):
                        self.batch_entries.append({
                            "pcb_id": "",
                            "mac": "",
                            "device_name": "",
                            "row_index": i
                        })

                    self.saved_batch_size = n
                    log_debug(f"Batch started with size: {n}")
                    self.show_batch_processing_page()
                except ValueError:
                    log_debug("start_batch: invalid integer")
                    messagebox.showerror("Error", "Please enter a valid number")

            ttk.Button(frame, text="Start Batch Processing", command=start_batch).pack(pady=10)
            ttk.Button(frame, text="Back", command=self.show_config_page).pack(pady=5)
        except Exception as e:
            log_debug(f"Exception in show_batch_setup_page: {e}")

    # -----------------------
    # Page: Batch processing
    # -----------------------
    def show_batch_processing_page(self):
        log_debug("show_batch_processing_page called")
        try:
            for w in self.main_frame.winfo_children():
                w.destroy()

            config_info = f"Version: {self.saved_version} | Code Version: {self.saved_code_version} | CSV: {os.path.basename(self.saved_csv_path) if self.saved_csv_path else ''} | Batch Size: {self.saved_batch_size}"
            ttk.Label(self.main_frame, text=config_info, font=("Arial", 10)).pack(anchor="w", pady=(0, 10))

            instructions = (
                "1. Enter PCB IDs for all boards first (use Tab to move between fields quickly)\n"
                "2. Connect serial port and start scanning\n"
                "3. Assign scanned devices to rows using 'MAC' buttons"
            )
            ttk.Label(self.main_frame, text=instructions, font=("Arial", 9), foreground="blue").pack(anchor="w", pady=(0, 10))

            # Serial controls
            serial_frame = ttk.Frame(self.main_frame)
            serial_frame.pack(fill="x", pady=(0, 10))

            ttk.Label(serial_frame, text="Serial Port:").pack(side="left", padx=(0, 5))
            port_combo = ttk.Combobox(serial_frame, width=15, values=self.get_available_ports())
            port_combo.pack(side="left", padx=(0, 10))

            ttk.Button(serial_frame, text="Refresh Ports", command=lambda: self.refresh_ports(port_combo)).pack(side="left", padx=(0, 10))

            connect_btn = ttk.Button(serial_frame, text="Connect", command=lambda: self.connect_serial(port_combo, connect_btn, disconnect_btn, start_btn, stop_btn))
            connect_btn.pack(side="left", padx=(0, 10))

            disconnect_btn = ttk.Button(serial_frame, text="Disconnect", state="disabled",
                                        command=lambda: self.disconnect_serial(connect_btn, disconnect_btn, start_btn, stop_btn))
            disconnect_btn.pack(side="left", padx=(0, 10))

            ttk.Label(serial_frame, text="Print Delay (s):").pack(side="left", padx=(20, 5))
            delay_entry = ttk.Entry(serial_frame, width=8)
            delay_entry.insert(0, str(PRINT_DELAY))
            delay_entry.pack(side="left", padx=(0, 10))

            def set_delay():
                log_debug("set_delay called")
                nonlocal delay_entry
                try:
                    val = float(delay_entry.get().strip())
                    if val >= 0:
                        global PRINT_DELAY
                        PRINT_DELAY = val
                        log_debug(f"PRINT_DELAY set to {PRINT_DELAY}")
                        self.set_status(f"Print delay set to {PRINT_DELAY} seconds")
                    else:
                        log_debug("set_delay: negative value entered")
                        messagebox.showerror("Error", "Delay must be a positive number")
                except ValueError:
                    log_debug("set_delay: invalid number")
                    messagebox.showerror("Error", "Please enter a valid number for delay")

            ttk.Button(serial_frame, text="Set Delay", command=set_delay).pack(side="left")

            # Filter
            filter_frame = ttk.Frame(self.main_frame)
            filter_frame.pack(fill="x", pady=(0, 10))
            ttk.Label(filter_frame, text="Filter by name (comma-separated, optional):").pack(anchor="w")
            filter_entry = ttk.Entry(filter_frame, width=50)
            filter_entry.pack(fill="x", pady=(5, 0))
            filter_entry.insert(0, "AT 52805")

            # Control buttons
            control_frame = ttk.Frame(self.main_frame)
            control_frame.pack(fill="x", pady=(5, 10))

            start_btn = ttk.Button(control_frame, text="Start", state="disabled", command=lambda: self.start_continuous_scan(filter_entry, start_btn, stop_btn))
            start_btn.pack(side="left", padx=(0, 10))

            ttk.Button(control_frame, text="Clear Scan", command=self.clear_scanned_devices).pack(side="left", padx=(0, 10))

            stop_btn = ttk.Button(control_frame, text="Stop", state="disabled", command=lambda: self.stop_continuous_scan(start_btn, stop_btn))
            stop_btn.pack(side="left", padx=(0, 10))

            ttk.Button(control_frame, text="Clear All IDs", command=self.clear_all_pcb_ids).pack(side="left", padx=(0, 10))

            ttk.Button(control_frame, text="Clear All MAC", command=self.clear_all_macs_and_scanned).pack(side="left", padx=(0, 10))

            add_all_btn = ttk.Button(control_frame, text="Add All", command=self.add_batch_to_csv)
            add_all_btn.pack(side="left", padx=(0, 10))
            ttk.Button(control_frame, text="Back to Setup", command=self.show_batch_setup_page).pack(side="left")

            ttk.Button(control_frame, text="View/Edit Data", command=self.show_csv_data).pack(pady=10)

            # Restore port selection and connect state if applicable
            if self.last_selected_port:
                try:
                    log_debug(f"Restoring last_selected_port: {self.last_selected_port}")
                    port_combo.set(self.last_selected_port)
                except Exception as e:
                    log_debug(f"Error setting port_combo value: {e}")

            if self.ser_connection and getattr(self.ser_connection, "is_open", False):
                log_debug("Serial connection already open; updating UI states")
                connect_btn.config(state="disabled")
                disconnect_btn.config(state="normal")
                start_btn.config(state="normal")
                stop_btn.config(state="disabled")
                self.set_status(f"Connected to {self.last_selected_port}")

            # Store widgets for later use
            self.widgets.port_combo = port_combo
            self.widgets.start_scan_button = start_btn
            self.widgets.stop_scan_button = stop_btn
            self.widgets.filter_entry = filter_entry
            self.widgets.add_all_button = add_all_btn
            self.widgets.delay_entry = delay_entry

            # Scrollable area for batch rows
            canvas = tk.Canvas(self.main_frame)
            scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=canvas.yview)
            scroll_frame = ttk.Frame(canvas)
            scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            self.result_frame = scroll_frame
            self.create_batch_entry_rows()
            self.update_batch_status()
        except Exception as e:
            log_debug(f"Exception in show_batch_processing_page: {e}")

    # -----------------------
    # Batch table UI
    # -----------------------
    def create_batch_entry_rows(self):
        log_debug("create_batch_entry_rows called")
        try:
            for w in self.result_frame.winfo_children():
                w.destroy()

            headers = ["#", "PCB ID", "MAC Address", "Device Name", "RSSI", "Status", "Actions"]
            for col, text in enumerate(headers):
                tk.Label(self.result_frame, text=text, font=("Arial", 9, "bold")).grid(row=0, column=col, padx=5, pady=2, sticky="w")

            for idx, entry in enumerate(self.batch_entries):
                row = idx + 1

                tk.Label(self.result_frame, text=str(idx + 1), font=("Arial", 9)).grid(row=row, column=0, padx=5, pady=1, sticky="w")

                pcb_var = tk.StringVar(value=entry.get("pcb_id", ""))
                pcb_entry = tk.Entry(self.result_frame, textvariable=pcb_var, width=15, font=("Arial", 9))
                pcb_entry.grid(row=row, column=1, padx=5, pady=1, sticky="w")

                # Tab navigation
                def tab_next(event, cur=idx):
                    if cur < len(self.batch_entries) - 1:
                        try:
                            self.batch_entries[cur + 1]["pcb_entry"].focus()
                        except Exception:
                            pass
                    return "break"

                def tab_prev(event, cur=idx):
                    if cur > 0:
                        try:
                            self.batch_entries[cur - 1]["pcb_entry"].focus()
                        except Exception:
                            pass
                    return "break"

                pcb_entry.bind('<Tab>', tab_next)
                pcb_entry.bind('<Shift-Tab>', tab_prev)

                def pcb_trace(*args, ri=idx, var=pcb_var):
                    try:
                        self.batch_entries[ri]["pcb_id"] = var.get()
                        self.update_batch_status()
                    except Exception:
                        log_debug("Error in pcb_trace")

                pcb_var.trace('w', pcb_trace)

                mac_var = tk.StringVar(value=entry.get("mac", ""))
                mac_label = tk.Label(self.result_frame, textvariable=mac_var, width=18, anchor="w", relief="solid", padx=2)
                mac_label.grid(row=row, column=2, padx=5, pady=1, sticky="w")

                name_var = tk.StringVar(value=entry.get("device_name", ""))
                name_label = tk.Label(self.result_frame, textvariable=name_var, width=12, anchor="w", relief="solid", padx=2)
                name_label.grid(row=row, column=3, padx=5, pady=1, sticky="w")

                rssi_var = tk.StringVar(value=entry.get("rssi", ""))
                rssi_label = tk.Label(self.result_frame, textvariable=rssi_var, width=6, anchor="w", relief="solid", padx=2)
                rssi_label.grid(row=row, column=4, padx=5, pady=1, sticky="w")

                status_lbl = tk.Label(self.result_frame, text="Empty", width=12, anchor="w", foreground="gray", font=("Arial", 9))
                status_lbl.grid(row=row, column=5, padx=5, pady=1, sticky="w")

                action_frame = ttk.Frame(self.result_frame)
                action_frame.grid(row=row, column=6, padx=2, pady=1, sticky="w")

                assign_btn = tk.Button(action_frame, text="MAC", font=("Arial", 8), width=8, command=lambda ri=idx: self.assign_next_device(ri))
                assign_btn.pack(side="left", padx=1)
                copy_btn = tk.Button(action_frame, text="Copy", font=("Arial", 8), width=5, command=lambda ri=idx: pyperclip.copy(self.batch_entries[ri].get("mac", "")))
                copy_btn.pack(side="left", padx=1)
                clear_mac_btn = tk.Button(action_frame, text="Clear", font=("Arial", 8), width=8, command=lambda ri=idx: self.clear_single_mac(ri))
                clear_mac_btn.pack(side="left", padx=1)
                add_btn = tk.Button(action_frame, text="Add", font=("Arial", 8), width=6, command=lambda ri=idx: self.add_single_to_csv(ri))
                add_btn.pack(side="left", padx=1)

                # Save ui refs inside entry dict
                entry.update({
                    "pcb_var": pcb_var,
                    "mac_var": mac_var,
                    "name_var": name_var,
                    "rssi_var": rssi_var,
                    "status_label": status_lbl,
                    "assign_btn": assign_btn,
                    "pcb_entry": pcb_entry,
                    "mac_label": mac_label,
                    "name_label": name_label,
                    "rssi_label": rssi_label
                })

                # update initial status
                self.update_entry_status(idx)
        except Exception as e:
            log_debug(f"Exception in create_batch_entry_rows: {e}")

    # -----------------------
    # Status helpers
    # -----------------------
    def set_status(self, text):
        try:
            log_debug(f"set_status: {text}")
            self.status_label.config(text=text)
        except Exception:
            log_debug("set_status failed")

    def update_entry_status(self, row_index):
        try:
            entry = self.batch_entries[row_index]
            pcb = entry.get("pcb_id", "").strip()
            mac = entry.get("mac", "").strip()
            label = entry.get("status_label")
            if not label:
                return

            if not pcb and not mac:
                label.config(text="Empty", foreground="gray")
            elif pcb and not mac:
                label.config(text="Ready for MAC", foreground="orange")
            elif not pcb and mac:
                label.config(text="Need PCB ID", foreground="red")
            else:
                label.config(text="Ready", foreground="green")
        except Exception:
            log_debug("update_entry_status error")

    def update_batch_status(self):
        try:
            ready = pcb_only = mac_only = empty = 0
            for entry in self.batch_entries:
                pcb = entry.get("pcb_id", "").strip()
                mac = entry.get("mac", "").strip()
                if pcb and mac:
                    ready += 1
                elif pcb and not mac:
                    pcb_only += 1
                elif not pcb and mac:
                    mac_only += 1
                else:
                    empty += 1
                # Attempt to update UI label if exists
                try:
                    self.update_entry_status(entry.get("row_index", 0))
                except Exception:
                    pass
            status_text = f"Batch: {ready} ready, {pcb_only} waiting for MAC, {mac_only} need PCB ID, {empty} empty"
            self.set_status(status_text)
            log_debug(f"update_batch_status: {status_text}")
        except Exception:
            log_debug("update_batch_status failed")

    # -----------------------
    # Assign / clear functions
    # -----------------------
    def assign_next_device(self, row_index):
        log_debug(f"assign_next_device called for row {row_index}")
        try:
            if not SCANNED_DEVICES:
                log_debug("assign_next_device: No devices scanned")
                messagebox.showinfo("No Devices", "No devices scanned yet. Start scanning to assign devices.")
                return

            latest = max(SCANNED_DEVICES.items(), key=lambda x: x[1]["serial_no"])
            mac, info = latest

            # check assigned elsewhere
            for i, e in enumerate(self.batch_entries):
                if e.get("mac") == mac and i != row_index:
                    log_debug(f"MAC {mac} already assigned to row {i+1}")
                    messagebox.showwarning("MAC Already Assigned", f"This MAC address is already assigned to row {i + 1}")
                    return

            self.batch_entries[row_index]["mac"] = mac
            self.batch_entries[row_index]["device_name"] = info.get("name", "")
            self.batch_entries[row_index]["mac_var"].set(mac)
            self.batch_entries[row_index]["name_var"].set(info.get("name", ""))
            self.batch_entries[row_index]["rssi_var"].set(str(info.get("rssi", "")))

            # remove from scanned to avoid reassign
            SCANNED_DEVICES.pop(mac, None)
            self.update_batch_status()
            self.set_status(f"Assigned {mac} to row {row_index + 1}")
            log_debug(f"Assigned MAC {mac} to row {row_index+1}")
        except Exception:
            log_debug("assign_next_device failed")

    def clear_single_mac(self, row_index):
        log_debug(f"clear_single_mac called for row {row_index}")
        try:
            e = self.batch_entries[row_index]
            e["mac"] = ""
            e["device_name"] = ""
            e["mac_var"].set("")
            e["name_var"].set("")
            e["rssi_var"].set("")
            self.update_batch_status()
        except Exception:
            log_debug("clear_single_mac failed")

    def clear_single_pcb(self, row_index):
        log_debug(f"clear_single_pcb called for row {row_index}")
        try:
            e = self.batch_entries[row_index]
            e["pcb_id"] = ""
            e["pcb_var"].set("")
            self.update_batch_status()
        except Exception:
            log_debug("clear_single_pcb failed")

    def clear_scanned_devices(self):
        log_debug("clear_scanned_devices called")
        global SCANNED_DEVICES
        SCANNED_DEVICES.clear()
        self.current_serial_number = 1
        self.set_status("Scanned devices cleared. Ready for new scans.")
        log_debug("SCANNED_DEVICES cleared")

    def clear_all_macs_and_scanned(self):
        log_debug("clear_all_macs_and_scanned called")
        if messagebox.askyesno("Confirm Clear", "Clear all MAC addresses and scanned devices?\n\nThis will remove all assigned MACs and clear the scan history."):
            for e in self.batch_entries:
                e["mac"] = ""
                e["device_name"] = ""
                try:
                    e["mac_var"].set("")
                    e["name_var"].set("")
                    e["rssi_var"].set("")
                except Exception:
                    pass
            self.clear_scanned_devices()
            self.update_batch_status()
            log_debug("All MACs and scanned devices cleared by user")

    def clear_all_macs(self):
        log_debug("clear_all_macs called")
        if messagebox.askyesno("Confirm Clear", "Clear all MAC addresses from all rows?"):
            for e in self.batch_entries:
                e["mac"] = ""
                e["device_name"] = ""
                try:
                    e["mac_var"].set("")
                    e["name_var"].set("")
                    e["rssi_var"].set("")
                except Exception:
                    pass
            self.update_batch_status()
            log_debug("All MACs cleared by user")

    def clear_all_pcb_ids(self):
        log_debug("clear_all_pcb_ids called")
        if messagebox.askyesno("Confirm Clear", "Clear all PCB IDs from all rows?"):
            for e in self.batch_entries:
                e["pcb_id"] = ""
                try:
                    e["pcb_var"].set("")
                except Exception:
                    pass
            self.update_batch_status()
            log_debug("All PCB IDs cleared by user")

    # -----------------------
    # CSV helpers
    # -----------------------
    def load_existing_csv_entries(self):
        log_debug("load_existing_csv_entries called")
        global csv_serial_counter
        self.csv_entries = []
        csv_serial_counter = 1
        if self.csv_file_path and os.path.exists(self.csv_file_path):
            try:
                with open(self.csv_file_path, "r", newline="") as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        if row.get('SL.No') and row['SL.No'].isdigit():
                            self.csv_entries.append(row)
                            current_num = int(row['SL.No'])
                            csv_serial_counter = max(csv_serial_counter, current_num + 1)
                log_debug(f"Loaded {len(self.csv_entries)} entries from CSV; next serial: {csv_serial_counter}")
            except Exception as exc:
                log_debug(f"Error reading CSV file: {exc}")
                # Try to recreate header safely
                try:
                    with open(self.csv_file_path, "w", newline="") as f:
                        writer = csv.writer(f)
                        writer.writerow(["SL.No", "VERSION", "CODE VERSION", "PCB ID", "MAC ID", "COMMENTS"])
                    log_debug("Recreated CSV header after read failure")
                except Exception as e:
                    log_debug(f"Failed recreating CSV header: {e}")
        else:
            log_debug("No CSV file path set or path does not exist")

    def add_to_csv(self, mac_address, pcb_id, device_name):
        log_debug(f"add_to_csv called for MAC={mac_address}, PCB={pcb_id}")
        global csv_serial_counter
        entry = {
            "SL.No": str(csv_serial_counter),
            "VERSION": self.version,
            "CODE VERSION": self.code_version,
            "PCB ID": pcb_id,
            "MAC ID": mac_address,
            "COMMENTS": device_name
        }
        self.csv_entries.append(entry)
        try:
            file_exists = os.path.exists(self.csv_file_path)
            with open(self.csv_file_path, "a", newline="") as f:
                fieldnames = ["SL.No", "VERSION", "CODE VERSION", "PCB ID", "MAC ID", "COMMENTS"]
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                if not file_exists or os.path.getsize(self.csv_file_path) == 0:
                    writer.writeheader()
                writer.writerow(entry)
            log_debug(f"Appended to CSV: SL.No={entry['SL.No']}, MAC={mac_address}, PCB={pcb_id}")
            csv_serial_counter += 1
            return entry["SL.No"]
        except Exception as exc:
            log_debug(f"Failed to append to CSV: {exc}")
            raise

    # -----------------------
    # Add single / batch operations
    # -----------------------
    def add_single_to_csv(self, row_index):
        log_debug(f"add_single_to_csv called for row {row_index}")
        try:
            e = self.batch_entries[row_index]
            pcb = e.get("pcb_id", "").strip()
            mac = e.get("mac", "").strip()
            name = e.get("device_name", "")

            if not pcb:
                log_debug("add_single_to_csv: PCB ID missing")
                messagebox.showerror("Error", "Please enter PCB ID for this row")
                return
            if not mac:
                log_debug("add_single_to_csv: MAC missing")
                messagebox.showerror("Error", "No MAC address assigned to this row")
                return

            sl = self.add_to_csv(mac, pcb, name)
            self.print_device([mac, pcb])

            # clear row after add
            e["pcb_id"] = ""
            e["mac"] = ""
            e["device_name"] = ""
            try:
                e["pcb_var"].set("")
                e["mac_var"].set("")
                e["name_var"].set("")
                e["rssi_var"].set("")
            except Exception:
                pass
            self.update_batch_status()
            log_debug(f"add_single_to_csv success: SL.No={sl}")
            messagebox.showinfo("Success", f"Added to CSV with S.No: {sl}")
        except Exception:
            log_debug("add_single_to_csv failed")
            messagebox.showerror("Error", "Failed to add row to CSV. See log for details.")

    def add_batch_to_csv(self):
        log_debug("add_batch_to_csv called")
        try:
            ready_entries = [e for e in self.batch_entries if e.get("pcb_id", "").strip() and e.get("mac", "").strip()]
            if not ready_entries:
                log_debug("add_batch_to_csv: No ready entries")
                messagebox.showinfo("No Entries", "No complete entries to add to CSV")
                return

            if not messagebox.askyesno("Confirm Add All", f"This will add {len(ready_entries)} device(s) to CSV and print labels.\n\nDo you want to continue?"):
                log_debug("User cancelled Add All")
                return

            # disable button
            try:
                if hasattr(self.widgets, "add_all_button"):
                    self.widgets.add_all_button.config(state="disabled")
            except Exception:
                pass
            self.set_status(f"Processing {len(ready_entries)} devices...")
            log_debug(f"Beginning batch add of {len(ready_entries)} devices")

            def worker():
                success = failed = 0
                for i, e in enumerate(ready_entries):
                    try:
                        self.add_to_csv(e.get("mac"), e.get("pcb_id"), e.get("device_name"))
                        self.print_device([e.get("mac"), e.get("pcb_id")])
                        success += 1
                        if i % 3 == 0:
                            self.root.after(0, lambda i=i, ln=len(ready_entries): self.set_status(f"Processed {i+1}/{ln} devices..."))
                    except Exception as exc:
                        log_debug(f"Error processing {e.get('mac')}: {exc}")
                        failed += 1

                # clear processed entries
                for e in ready_entries:
                    e["pcb_id"] = ""
                    e["mac"] = ""
                    e["device_name"] = ""
                    try:
                        e["pcb_var"].set("")
                        e["mac_var"].set("")
                        e["name_var"].set("")
                        e["rssi_var"].set("")
                    except Exception:
                        pass

                self.root.after(0, lambda: self.finish_batch_process(success, failed))

            threading.Thread(target=worker, daemon=True).start()
        except Exception:
            log_debug("add_batch_to_csv failed")

    def finish_batch_process(self, success_count, failed_count):
        log_debug(f"finish_batch_process called: success={success_count}, failed={failed_count}")
        try:
            try:
                if hasattr(self.widgets, "add_all_button"):
                    self.widgets.add_all_button.config(state="normal")
            except Exception:
                pass
            self.update_batch_status()
            if failed_count == 0:
                messagebox.showinfo("Success", f"Successfully added and printed {success_count} device(s).")
                log_debug(f"Batch completed successfully: {success_count} devices")
            else:
                messagebox.showwarning("Partial Success", f"Added {success_count} device(s) successfully.\nFailed to process {failed_count} device(s).")
                log_debug(f"Batch completed with failures: success={success_count}, failed={failed_count}")
            self.set_status(f"Batch completed. Added {success_count} device(s).")
        except Exception:
            log_debug("finish_batch_process failed")

    def show_csv_data(self):
        log_debug("show_csv_data called")
        if not self.csv_file_path or not os.path.exists(self.csv_file_path):
            log_debug("show_csv_data: No valid CSV file selected.")
            messagebox.showerror("Error", "No valid CSV file selected.")
            return

        try:
            win = tk.Toplevel(self.root)
            win.title("CSV Data Viewer")
            win.geometry("850x400")

            frame = ttk.Frame(win, padding=10)
            frame.pack(fill="both", expand=True)

            cols = ["SL.No", "VERSION", "CODE VERSION", "ID", "MAC ID", "COMMENTS"]
            tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")

            for c in cols:
                tree.heading(c, text=c)
                if c == "SL.No":
                    tree.column(c, width=60, anchor="center")
                elif c in ("VERSION", "ID"):
                    tree.column(c, width=100, anchor="center")
                elif c in ("CODE VERSION",):
                    tree.column(c, width=200, anchor="center")
                elif c == "MAC ID":
                    tree.column(c, width=150, anchor="center")
                else:
                    tree.column(c, width=100, anchor="center")

            yscroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=yscroll.set)
            tree.pack(side="left", fill="both", expand=True)
            yscroll.pack(side="right", fill="y")

            btn_frame = ttk.Frame(win)
            btn_frame.pack(fill="x", pady=5)
            ttk.Button(btn_frame, text="Refresh", command=lambda: load_data()).pack(side="left", padx=5)
            ttk.Button(btn_frame, text="Delete Selected Row", command=lambda: delete_selected()).pack(side="left", padx=5)
            ttk.Button(btn_frame, text="Close", command=win.destroy).pack(side="right", padx=5)

            def load_data():
                log_debug("CSV viewer load_data called")
                for r in tree.get_children():
                    tree.delete(r)
                try:
                    with open(self.csv_file_path, "r", newline="") as f:
                        reader = csv.DictReader(f)
                        for row in reader:
                            vals = [row.get(c, "") for c in cols]
                            tree.insert("", "end", values=vals)
                    log_debug("CSV viewer loaded data")
                except Exception as e:
                    log_debug(f"CSV viewer failed to load data: {e}")
                    messagebox.showerror("Error", f"Failed to read CSV: {e}")

            def reindex_and_save(rows):
                log_debug("reindex_and_save called")
                for i, r in enumerate(rows, start=1):
                    r["SL.No"] = str(i)
                with open(self.csv_file_path, "w", newline="") as f:
                    writer = csv.DictWriter(f, fieldnames=cols)
                    writer.writeheader()
                    writer.writerows(rows)
                log_debug("CSV reindexed and saved")

            def delete_selected():
                log_debug("delete_selected called in CSV viewer")
                selected = tree.selection()
                if not selected:
                    log_debug("delete_selected: no selection")
                    messagebox.showinfo("No Selection", "Please select a row to delete.")
                    return

                vals = tree.item(selected[0], "values")
                slno = vals[0]
                if not messagebox.askyesno("Confirm Delete", f"Delete row with SL.No {slno}?"):
                    log_debug("delete_selected: user canceled deletion")
                    return

                try:
                    rows = []
                    with open(self.csv_file_path, "r", newline="") as f:
                        reader = csv.DictReader(f)
                        for row in reader:
                            if row.get("SL.No") != slno:
                                rows.append(row)
                    reindex_and_save(rows)
                    messagebox.showinfo("Deleted", f"Row with SL.No {slno} deleted and CSV reindexed.")
                    load_data()
                    self.load_existing_csv_entries()
                    log_debug(f"Deleted CSV row SL.No={slno}")
                except Exception as e:
                    log_debug(f"delete_selected: failed to delete row {slno}: {e}")
                    messagebox.showerror("Error", f"Failed to delete row: {e}")

            load_data()
        except Exception:
            log_debug("show_csv_data failed")

    # -----------------------
    # Serial communication
    # -----------------------
    def get_available_ports(self):
        try:
            ports = [p.device for p in serial.tools.list_ports.comports()]
            log_debug(f"get_available_ports: found {ports}")
            return ports
        except Exception:
            log_debug("get_available_ports failed")
            return []

    def refresh_ports(self, combo_widget=None):
        log_debug("refresh_ports called")
        vals = self.get_available_ports()
        try:
            if combo_widget is None:
                if hasattr(self.widgets, "port_combo"):
                    self.widgets.port_combo["values"] = vals
            else:
                combo_widget["values"] = vals
            log_debug("Serial ports refreshed")
        except Exception:
            log_debug("refresh_ports failed to update widget")

    def connect_serial(self, port_combo, connect_btn, disconnect_btn, start_btn, stop_btn):
        log_debug("connect_serial called")
        port = port_combo.get()
        self.last_selected_port = port
        if not port:
            log_debug("connect_serial: no port selected")
            messagebox.showerror("Error", "Please select a serial port")
            return
        try:
            log_debug(f"Attempting to open serial port: {port}")
            self.ser_connection = serial.Serial(port, 115200, timeout=1)
            log_debug(f"Serial port {port} opened")
            self.set_status(f"Connected to {port}")
            try:
                connect_btn.config(state="disabled")
                disconnect_btn.config(state="normal")
                start_btn.config(state="normal")
                stop_btn.config(state="disabled")
            except Exception:
                pass
            self.is_scanning = True
            # start reader thread
            self.serial_thread = threading.Thread(target=self.read_serial_data, daemon=True)
            self.serial_thread.start()
            log_debug("Serial reader thread started")
        except Exception as exc:
            log_debug(f"connect_serial: Failed to connect: {exc}")
            messagebox.showerror("Connection Error", f"Failed to connect: {exc}")

    def disconnect_serial(self, connect_btn, disconnect_btn, start_btn, stop_btn):
        log_debug("disconnect_serial called")
        self.is_scanning = False
        try:
            if self.ser_connection and getattr(self.ser_connection, "is_open", False):
                try:
                    self.send_to_arduino("STOP")
                except Exception:
                    log_debug("Failed sending STOP to Arduino during disconnect (ignored)")
                try:
                    self.ser_connection.close()
                    log_debug("Serial connection closed")
                except Exception:
                    log_debug("Error closing serial connection")
        except Exception:
            log_debug("disconnect_serial encountered unexpected error")

        self.ser_connection = None
        try:
            connect_btn.config(state="normal")
            disconnect_btn.config(state="disabled")
            start_btn.config(state="disabled")
            stop_btn.config(state="disabled")
        except Exception:
            pass
        self.set_status("Disconnected")

    def send_to_arduino(self, command):
        log_debug(f"send_to_arduino called with command: {command}")
        if self.ser_connection and getattr(self.ser_connection, "is_open", False):
            try:
                self.ser_connection.write(f"{command}\n".encode())
                log_debug(f"Sent to Arduino: {command}")
            except Exception as exc:
                log_debug(f"Error sending to Arduino: {exc}")
        else:
            log_debug("send_to_arduino: serial connection not open")

    def start_continuous_scan(self, filter_entry, start_btn, stop_btn):
        log_debug("start_continuous_scan called")
        self.is_scanning = True
        try:
            start_btn.config(state="disabled")
            stop_btn.config(state="normal")
        except Exception:
            pass
        self.set_status("Starting scan...")
        if self.ser_connection and getattr(self.ser_connection, "is_open", False):
            try:
                self.ser_connection.reset_input_buffer()
                log_debug("Serial input buffer reset for scan")
            except Exception:
                log_debug("Failed resetting serial input buffer")
        self.send_to_arduino("START")

    def stop_continuous_scan(self, start_btn, stop_btn):
        log_debug("stop_continuous_scan called")
        self.is_scanning = False
        try:
            start_btn.config(state="normal")
            stop_btn.config(state="disabled")
        except Exception:
            pass
        self.set_status("Stopping scan...")
        self.send_to_arduino("STOP")
        if self.ser_connection and getattr(self.ser_connection, "is_open", False):
            try:
                self.ser_connection.reset_input_buffer()
            except Exception:
                log_debug("Failed resetting serial input buffer on stop")

    def read_serial_data(self):
        log_debug("read_serial_data thread started")
        while self.ser_connection and getattr(self.ser_connection, "is_open", False):
            try:
                if getattr(self.ser_connection, "in_waiting", 0):
                    line = self.ser_connection.readline().decode('utf-8', errors='ignore').strip()
                    if line:
                        log_debug(f"Serial received: {line}")
                        self.process_serial_data(line)
                time.sleep(0.01)
            except Exception as exc:
                log_debug(f"Serial read error: {exc}")
                time.sleep(0.5)
        log_debug("read_serial_data thread exiting (serial closed)")

    def process_serial_data(self, data):
        try:
            log_debug(f"process_serial_data called with data: {data}")
            # expected special markers
            if data == "SCANNING_STARTED":
                self.set_status("Scanning started - receiving data...")
                log_debug("Received SCANNING_STARTED")
                return
            if data == "SCANNING_STOPPED":
                self.set_status("Scanning stopped")
                log_debug("Received SCANNING_STOPPED")
                return
            if data == "SINGLE_SCAN":
                self.set_status("Single scan performed")
                log_debug("Received SINGLE_SCAN")
                return

            # data format: START,<device_name>,<mac>,<rssi>,END
            if data.startswith("START,") and data.endswith(",END"):
                payload = data[6:-4]
                parts = payload.split(',')
                if len(parts) == 3:
                    device_name, mac, rssi_str = parts
                    mac = mac.upper()

                    # apply filter if present
                    filter_text = getattr(self.widgets, "filter_entry", None)
                    if filter_text:
                        try:
                            flt = filter_text.get().strip()
                        except Exception:
                            flt = ""
                        if flt:
                            filters = [f.strip().lower() for f in flt.split(',') if f.strip()]
                            if not any(fn in device_name.lower() for fn in filters):
                                log_debug(f"Filtered out device {device_name} by filter {filters}")
                                return

                    if mac in IGNORED_DEVICES:
                        log_debug(f"Ignoring device {mac} (in IGNORED_DEVICES)")
                        return

                    try:
                        rssi = int(rssi_str)
                    except ValueError:
                        log_debug(f"Invalid RSSI value received: {rssi_str}")
                        return

                    # add or update scanned device
                    if mac not in SCANNED_DEVICES:
                        SCANNED_DEVICES[mac] = {"name": device_name, "rssi": rssi, "serial_no": self.current_serial_number}
                        self.current_serial_number += 1
                        log_debug(f"New device scanned: {device_name} {mac} rssi={rssi}")
                    else:
                        SCANNED_DEVICES[mac]["rssi"] = rssi
                        SCANNED_DEVICES[mac]["name"] = device_name
                        log_debug(f"Updated scanned device: {device_name} {mac} rssi={rssi}")

                    self.set_status(f"Scanned: {device_name} - {mac}")
                else:
                    log_debug(f"process_serial_data: unexpected START payload parts: {parts}")
            else:
                log_debug("process_serial_data: data did not match expected patterns")
        except Exception:
            log_debug("process_serial_data failed")

    # -----------------------
    # Printing / Bartender
    # -----------------------
    def print_with_bartender(self, btw_path, value):
        log_debug(f"print_with_bartender called with btw_path={btw_path} value={value}")
        try:
            pythoncom.CoInitialize()
            bt = win32com.client.Dispatch("Bartender.Application")
            bt.Visible = False
            fmt = bt.Formats.Open(btw_path, False, "")
            try:
                fmt.SetNamedSubStringValue("MAC", value[0])
                fmt.SetNamedSubStringValue("ID",  value[1])
            except Exception:
                log_debug("print_with_bartender: setting sub strings failed")
            fmt.PrintOut(False, False)
            fmt.Close(1)
            bt.Quit()
            log_debug(f"print_with_bartender sent print for {value}")
        except Exception as exc:
            log_debug(f"print_with_bartender error: {exc}")

    def _print_worker(self):
        log_debug("[BT] Print worker thread started")
        while True:
            try:
                mac, pcb = self.print_queue.get()
                log_debug(f"[BT] Worker processing: MAC={mac}, PCB={pcb}")
                
                if not self._bt_printer:
                    log_debug("[BT] Creating new BartenderPrinter instance")
                    self._bt_printer = BartenderPrinter(self.btw_file_path)
                
                if self._bt_printer and self._bt_printer.format:
                    success = self._bt_printer.print_label(mac, pcb)
                    if success:
                        log_debug(f"[BT] Print job completed successfully for {mac}")
                    else:
                        log_debug(f"[BT] Print job failed for {mac}")
                else:
                    log_debug("[BT] No valid Bartender printer available")
                    messagebox.showerror("Print Error", "Bartender is not properly initialized")
                
                log_debug(f"[BT] Sleeping for {PRINT_DELAY} seconds")
                time.sleep(PRINT_DELAY)
                
            except Exception as e:
                log_debug(f"[BT] Worker error: {e}")
                messagebox.showerror("Print Error", f"Print worker error:\n{e}")
            finally:
                try:
                    self.print_queue.task_done()
                except Exception as e:
                    log_debug(f"[BT] Error calling task_done: {e}")

    def print_device(self, value):
        log_debug(f"print_device called with value: {value}")
        if not self.btw_file_path:
            log_debug("print_device called but no BTW file selected")
            messagebox.showwarning("No File Selected", "Please select a Bartender label file first.")
            return
        try:
            mac, pcb = value
            self.print_queue.put((mac, pcb))
            self.set_status(f"Queued print for {mac} / {pcb}")
            log_debug(f"Queued print for MAC={mac}, PCB={pcb}")
        except Exception:
            log_debug("print_device failed")

    def on_close(self):
        log_debug("[APP] on_close called - cleaning up")
        try:
            if hasattr(self, "_bt_printer") and self._bt_printer:
                try:
                    self._bt_printer.close()
                except Exception:
                    log_debug("Error closing _bt_printer during on_close")
            # Close serial if open
            try:
                if self.ser_connection and getattr(self.ser_connection, "is_open", False):
                    try:
                        self.send_to_arduino("STOP")
                    except Exception:
                        pass
                    try:
                        self.ser_connection.close()
                        log_debug("Serial connection closed on app exit")
                    except Exception:
                        log_debug("Failed to close serial on exit")
            except Exception:
                log_debug("Error handling serial connection on close")
        except Exception:
            log_debug("Unexpected error in on_close cleanup")
        log_debug("[APP] Cleanup complete. Exiting.")
        try:
            self.root.destroy()
        except Exception:
            log_debug("root.destroy failed in on_close")

def clear_print_queue(printer_name=None):
    log_debug("clear_print_queue called")
    if not printer_name:
        try:
            printer_name = win32print.GetDefaultPrinter()
        except Exception:
            log_debug("Could not get default printer")
            printer_name = None
    if not printer_name:
        log_debug("No printer name available for clear_print_queue")
        return
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            jobs = win32print.EnumJobs(hPrinter, 0, -1, 1)
            for job in jobs:
                try:
                    win32print.SetJob(hPrinter, job['JobId'], 0, None, 2)
                    log_debug(f"Deleted print job {job['JobId']}")
                except Exception:
                    log_debug(f"Failed to delete print job {job.get('JobId')}")
        finally:
            win32print.ClosePrinter(hPrinter)
            log_debug("Closed printer handle after clearing jobs")
    except Exception:
        log_debug("clear_print_queue encountered error")

# -----------------------
# Main
# -----------------------
def main():
    with open(DEBUG_LOG, "w") as f:
        f.write(f"STM32 Programmer Debug Log - {datetime.now()}\n\n")
        log_debug("=== App starting ===")
    try:
        clear_print_queue()  # optional
    except Exception as e:
        log_debug(f"Error clearing print queue: {e}")

    root = tk.Tk()
    app = App(root)

    try:
        root.mainloop()
    except Exception as e:
        log_debug(f"Runtime error: {e}")
    finally:
        log_debug("=== App exiting ===")

if __name__ == "__main__":
    main()
