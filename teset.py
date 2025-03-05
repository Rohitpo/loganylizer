import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import serial
import serial.tools.list_ports
import os
import threading
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
from datetime import datetime

# Hardcoded keywords for highlighting
HIGHLIGHT_KEYWORDS = ["ERROR", "FAIL", "WARNING"]
user_keywords = []
ser_list = []  # List to hold multiple serial objects
log_frames = []  # List to store log display frames

def get_usb_ports():
    """Detect available USB ports."""
    ports = serial.tools.list_ports.comports()
    return [port.device for port in ports]

def save_log(data, filename):
    """Save logs automatically with timestamped filename."""
    with open(filename, 'w') as file:
        file.writelines(data)
    return filename

def generate_excel(log_file):
    """Convert log file into an Excel table."""
    logs = []
    with open(log_file, 'r') as file:
        for line in file:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            highlighted = any(keyword in line for keyword in HIGHLIGHT_KEYWORDS + user_keywords)
            logs.append([timestamp, line.strip(), "Yes" if highlighted else "No"])
    df = pd.DataFrame(logs, columns=["Timestamp", "Log Entry", "Highlighted"])
    excel_file = log_file.replace(".txt", ".xlsx")
    df.to_excel(excel_file, index=False)

def send_command(ser, command, log_display):
    """Send command to the serial device and print response in the log display."""
    if ser:
        ser.write((command + '\n').encode())
        log_display.insert(tk.END, f"Sent: {command}\n")
        log_display.see(tk.END)

def log_usb_data(ser, filename, log_display):
    """Read serial data and update GUI."""
    logs = []
    while True:
        try:
            line = ser.readline().decode(errors='ignore')
            if line:
                log_display.insert(tk.END, line)
                log_display.see(tk.END)
                logs.append(line)
                
                if any(k in line for k in HIGHLIGHT_KEYWORDS + user_keywords):
                    log_display.tag_add("highlight", f"{tk.END}-1l", tk.END)
                    log_display.tag_config("highlight", foreground="red")
            
            if len(logs) >= 60:
                save_log(logs, filename)
                generate_excel(filename)
                logs.clear()
        except Exception as e:
            break

def add_configuration():
    """Add another serial port configuration in a new window."""
    new_window = tk.Toplevel(root)
    new_window.title("New Configuration")
    
    port_var = tk.StringVar()
    baudrate_var = tk.StringVar(value="9600")
    filename_var = tk.StringVar()
    command_var = tk.StringVar()
    log_display = scrolledtext.ScrolledText(new_window, wrap=tk.WORD, height=15, width=80)
    log_display.pack(pady=5)
    
    ports = get_usb_ports()
    ttk.Label(new_window, text="Select Port:").pack()
    port_dropdown = ttk.Combobox(new_window, textvariable=port_var, values=ports)
    port_dropdown.pack()
    ttk.Label(new_window, text="Baud Rate:").pack()
    baudrate_entry = ttk.Entry(new_window, textvariable=baudrate_var)
    baudrate_entry.pack()
    
    ttk.Label(new_window, text="Save File:").pack()
    folder_button = ttk.Button(new_window, text="Browse", command=lambda: filename_var.set(filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])) )
    folder_button.pack()
    
    start_button = ttk.Button(new_window, text="Start Logging", command=lambda: start_logging(port_var.get(), int(baudrate_var.get()), filename_var.get(), log_display))
    start_button.pack()
    
    ttk.Label(new_window, text="Enter Command:").pack()
    command_entry = ttk.Entry(new_window, textvariable=command_var)
    command_entry.pack()
    command_button = ttk.Button(new_window, text="Send", command=lambda: send_command(ser_list[-1], command_var.get(), log_display))
    command_button.pack()
    
    ttk.Label(new_window, text="Find Keyword:").pack()
    find_entry = ttk.Entry(new_window)
    find_entry.pack()
    find_button = ttk.Button(new_window, text="Find", command=lambda: highlight_keyword(log_display, find_entry.get()))
    find_button.pack()

def highlight_keyword(log_display, keyword):
    """Highlight the keyword in logs."""
    log_display.tag_remove("found", "1.0", tk.END)
    if keyword:
        start_pos = "1.0"
        while True:
            start_pos = log_display.search(keyword, start_pos, stopindex=tk.END)
            if not start_pos:
                break
            end_pos = f"{start_pos}+{len(keyword)}c"
            log_display.tag_add("found", start_pos, end_pos)
            log_display.tag_config("found", background="yellow")
            start_pos = end_pos

root = tk.Tk()
root.title("Serial Logger GUI")
root.geometry("900x600")

top_frame = ttk.Frame(root)
top_frame.pack(fill=tk.X)

add_button = ttk.Button(top_frame, text="+ Add Configuration", command=add_configuration)
add_button.pack(side=tk.LEFT, padx=5)
reset_button = ttk.Button(top_frame, text="Reset", command=lambda: reset_logs(messagebox.askquestion("Reset", "Port Reset, Full Reset, or Cancel?")))
reset_button.pack(side=tk.LEFT, padx=5)

root.mainloop()