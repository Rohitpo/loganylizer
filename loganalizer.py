import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import random
import xlsxwriter  # Importing xlsxwriter
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

class LogAnalyzerApp:
    def __init__(self, parent):
        self.parent = parent
        self.setup_ui()
        self.df = pd.DataFrame(columns=['Time', 'Log Entry'])
        self.keywords = []
        self.keyword_colors = {}
        self.file_path = ""
        self.output_folder = ""

    def setup_ui(self):
        self.frame = ttk.Frame(self.parent, padding="10")
        self.frame.pack(fill=tk.BOTH, expand=True)

        # Top Controls
        self.top_frame = ttk.Frame(self.frame)
        self.top_frame.pack(fill=tk.X)

        self.upload_file_button = ttk.Button(self.top_frame, text="Upload Log File", command=self.upload_log_file)
        self.upload_file_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.uploaded_file_label = ttk.Label(self.top_frame, text="No file selected")
        self.uploaded_file_label.pack(side=tk.LEFT, padx=5, pady=5)

        # Keyword Entry
        self.keyword_entry = ttk.Entry(self.top_frame, width=20)
        self.keyword_entry.pack(side=tk.LEFT, padx=5)

        self.add_keyword_button = ttk.Button(self.top_frame, text="Add Keyword", command=self.add_keyword)
        self.add_keyword_button.pack(side=tk.LEFT, padx=5)

        # Keyword Listbox
        self.keywords_listbox = tk.Listbox(self.top_frame, height=5, width=30)
        self.keywords_listbox.pack(side=tk.LEFT, padx=5)
        self.keywords_listbox.bind("<<ListboxSelect>>", self.on_keyword_select)

        self.delete_keyword_button = ttk.Button(self.top_frame, text="Delete Keyword", command=self.delete_keyword)
        self.delete_keyword_button.pack(side=tk.LEFT, padx=5)

        # Output Folder Selection
        self.output_button = ttk.Button(self.top_frame, text="Select Output Folder", command=self.select_output_folder)
        self.output_button.pack(side=tk.LEFT, padx=5)
        self.output_folder_label = ttk.Label(self.top_frame, text="No folder selected")
        self.output_folder_label.pack(side=tk.LEFT, padx=5)

        # Analyze Logs Button
        self.analyze_button = ttk.Button(self.top_frame, text="Analyze Logs", command=self.analyze_logs)
        self.analyze_button.pack(side=tk.LEFT, padx=5)

        # Log Display
        self.logs_frame = ttk.Frame(self.frame)
        self.logs_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.logs_text = tk.Text(self.logs_frame, height=15, wrap=tk.WORD)
        self.logs_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = ttk.Scrollbar(self.logs_frame, orient=tk.VERTICAL, command=self.logs_text.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.logs_text.config(yscrollcommand=self.scrollbar.set)

        # Graph Area
        self.graph_area = ttk.Frame(self.frame)
        self.graph_area.pack(fill=tk.BOTH, expand=True)

    def upload_log_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Log Files", "*.txt *.log"), ("All Files", "*.*")])
        if file_path:
            self.file_path = file_path
            self.uploaded_file_label.config(text=os.path.basename(file_path))
            self.load_log_file()

    def load_log_file(self):
        if not self.file_path:
            return

        new_df = pd.DataFrame(columns=['Time', 'Log Entry'])
        log_text = ""

        try:
            with open(self.file_path, 'r', encoding='utf-8') as file:
                for line in file:
                    parts = line.strip().split(maxsplit=1)
                    if len(parts) == 2:
                        try:
                            timestamp = pd.to_datetime(parts[0], format='%H:%M:%S.%f')
                            new_df = pd.concat([new_df, pd.DataFrame({'Time': [timestamp], 'Log Entry': [parts[1]]})],
                                               ignore_index=True)
                            log_text += f"{timestamp}: {parts[1]}\n"
                        except ValueError:
                            continue

            self.df = new_df
            self.logs_text.delete(1.0, tk.END)
            self.logs_text.insert(tk.END, log_text)
            messagebox.showinfo("Success", "Log file loaded successfully.")
            self.plot_graph()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

    def add_keyword(self):
        keyword = self.keyword_entry.get().strip()
        if keyword and keyword not in self.keywords:
            self.keywords.append(keyword)
            self.keywords_listbox.insert(tk.END, keyword)
            self.keyword_entry.delete(0, tk.END)
            self.plot_graph()
        else:
            messagebox.showwarning("Warning", "Enter a unique keyword.")

    def on_keyword_select(self, event):
        selected_idx = self.keywords_listbox.curselection()
        if selected_idx:
            selected_keyword = self.keywords_listbox.get(selected_idx)
            self.keyword_entry.delete(0, tk.END)
            self.keyword_entry.insert(0, selected_keyword)

    def delete_keyword(self):
        selected_idx = self.keywords_listbox.curselection()
        if selected_idx:
            keyword = self.keywords_listbox.get(selected_idx)
            self.keywords.remove(keyword)
            self.keywords_listbox.delete(selected_idx)
            self.plot_graph()

    def select_output_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder = folder_path
            self.output_folder_label.config(text=folder_path)

    def analyze_logs(self):
        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        output_file = os.path.join(self.output_folder, f"Log_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                self.df.to_excel(writer, sheet_name='All Logs', index=False)

                # Summary Sheet
                summary = pd.DataFrame({'Keyword': self.keywords,
                                        'Occurrences': [self.df['Log Entry'].str.contains(k, case=False).sum() for k in self.keywords]})
                summary.to_excel(writer, sheet_name='Summary', index=False)

                for keyword in self.keywords:
                    keyword_data = []
                    indices = self.df[self.df['Log Entry'].str.contains(keyword, case=False)].index

                    for idx in indices:
                        start_idx = max(0, idx - 2)
                        end_idx = min(len(self.df) - 1, idx + 2)
                        keyword_data.append(self.df.iloc[start_idx:end_idx + 1])
                        keyword_data.append(pd.DataFrame([['', '']], columns=['Time', 'Log Entry']))  # 2-line gap

                    if keyword_data:
                        pd.concat(keyword_data).to_excel(writer, sheet_name=keyword[:30], index=False)

            messagebox.showinfo("Analysis Complete", f"Report saved to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving report: {e}")

    def plot_graph(self):
        if hasattr(self, 'canvas'):
            self.canvas.get_tk_widget().destroy()

        fig, ax = plt.subplots(figsize=(10, 5))
        for keyword in self.keywords:
            occurrences = self.df[self.df['Log Entry'].str.contains(keyword, case=False)]
            ax.plot(occurrences['Time'], [1] * len(occurrences), marker='o', linestyle='', label=keyword)

        ax.set_title('Keyword Occurrences Over Time')
        ax.set_xlabel('Time')
        ax.set_ylabel('Occurrences')
        ax.legend()
        ax.grid(True)

        self.canvas = FigureCanvasTkAgg(fig, self.graph_area)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

# Run the Application
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Log Analyzer")
    root.geometry("1000x700")

    log_analyzer_app = LogAnalyzerApp(root)

    root.mainloop()
