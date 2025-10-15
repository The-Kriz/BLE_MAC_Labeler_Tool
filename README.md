# BLE_MAC_Labeler_Tool

A factory production-line tool for scanning Bluetooth Low Energy (BLE) devices, capturing their MAC addresses, printing labels, and logging production data to CSV files.

---

## 📦 Overview

**BLE_MAC_Labeler_Tool** automates the process of identifying BLE modules during production testing.  
It connects to a serial-based BLE scanner (e.g., an ESP32 or Arduino-based scanner), receives detected BLE MAC addresses, pairs them with PCB IDs, and then:

- Prints labels via **Bartender Automation**
- Logs results (MAC, PCB ID, firmware version, etc.) to a CSV file for traceability
- Supports **batch scanning** and **multi-device production workflows**

This tool is designed for **factory floor environments** where large batches of BLE-based PCBs are tested, labeled, and logged efficiently.

---

## 🚀 Features

✅ **BLE MAC Scanning**
- Detects BLE devices via serial connection (from Arduino/ESP-based scanner)
- Filters device names (e.g., “AT52805”) for relevant modules only  
- Displays scanned MAC addresses and signal strength (RSSI)

✅ **Batch Labeling**
- Allows batch setup (number of boards to process)
- Lets you input PCB IDs and assign scanned MAC addresses per board
- Automatically manages serial connection, device assignment, and status tracking

✅ **Automated Label Printing**
- Integrates with **Bartender Label Software**
- Prints labels containing PCB ID and MAC address
- Queues print jobs and manages timing between prints

✅ **CSV Data Logging**
- Appends each entry with:
  - Serial number
  - Version / Code version
  - PCB ID
  - MAC address
  - Comments (e.g., device name)
- Auto-creates or validates CSV structure for consistent data logging

✅ **Production-Friendly UI**
- Built with Tkinter (Python GUI)
- Separate pages for configuration, batch setup, and live batch processing
- Non-blocking serial communication and print queue handling
- Option to clear print queue automatically on startup

✅ **Persistent Session Memory**
- Remembers selected BTW file, CSV path, serial port, and batch size between screens
- Keeps serial connection active even when navigating between pages

✅ **Debug Logging**
- Creates a detailed `BLE_6v2_debug_log.txt` log on every startup
- Overwrites old logs to keep clean session data
- Captures all events, serial communication, and Bartender interactions
---

## 🧩 Arduino BLE Scanner

The Scanner/Scanner.ino file contains the code for your serial BLE scanner.
It performs BLE scanning and outputs data in the following format:
```bash
START,<device_name>,<mac>,<rssi>,END
```
Example:
```
START,AT52805,AA:BB:CC:DD:EE:FF,-62,END
```
## ⚙️ Requirements

### Windows Environment
- Python 3.8 or newer  
- Bartender Label Software (for label printing)  
- Arduino BLE scanner connected over USB (or any BLE-to-serial device)

### Python Dependencies
Install required packages before running the script:
```bash
pip install pyserial pyperclip pywin32
```

🧰 Additional Utilities

Print Queue Cleaner: Automatically clears Windows print queue on app startup

Auto Debug Logs: Each session generates a new BLE_6v2_debug_log.txt file for troubleshooting

Error Handling: All exceptions and print errors are logged in the debug file


