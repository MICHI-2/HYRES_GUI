import tkinter as tk
from tkinter import ttk, scrolledtext
import json
import os
import subprocess
import threading
import to_xlsx

class HYRESrunnerAPP:
    def __init__(self, root):
        self.root = root
        self.root.title("HYRES GUI")
        
        self.entries = {
            "inout file name": {
                "cstar csv file name": "cstar csv file name",
                "gamma csv file name": "gamma csv file name"
            },
            "oxidizer": {
                "Initial tank pressure [Nm^-2]": "初期タンク圧力 [MPa]",
                "Final tank pressure [Nm^-2]": "燃焼終了時タンク圧力 [MPa]",
                "Oxidizer filling volume [m^3]": "酸化剤充填量 [cc]",
                "Oxidizer density [kgm^-3]": "酸化剤密度 [kg/m\u00B3]"
            },
            "fuel": {
                "Fuel density [kgm^-3]": "燃料密度 [kg/m\u00B3]",
                "Fuel length [m]": "燃料長さ [mm]",
                "Initial port diameter [m]": "初期ポート直径 [mm]",
                "Fuel outer diameter [m]": "燃料外径 [mm]",
                "Fuel port number [-]": "燃料ポート数 [-]"
            },
            "Oxidizer flow characteristics": {
                "Orifice diameter [m]": "オリフィス径 [mm]",
                "Flow coefficient [-]": "オリフィス流量係数 [-]"
            },
            "Combustion characteristics": {
                "Oxidizer mass flux coefficient [m^3kg^-1]": "酸化剤流束係数 [m\u00B3/kg]",
                "Oxidizer mass flux exponent [-]": "酸化剤流束指数 [-]",
                "C-star efficiency [-]": "C*効率 [-]"
            },
            "Nozzle characteristics": {
                "Initial nozzle throat diameter [m]": "初期ノズルスロート径 [mm]",
                "Nozzle exit diameter [m]": "ノズル出口径 [mm]",
                "Nozzle exit half angle [deg]": "ノズル半頂角 [deg]",
                "Nozzle erosion speed [ms^-1]": "ノズルスロートエロージョン速度 [mm/s]"
            },
            "Environment": {
                "Back pressure [Nm^-2]": "背圧 [MPa]"
            }
        }

        self.entry_data = {}
        self.pressure_keys = {
            "Initial tank pressure [Nm^-2]",
            "Final tank pressure [Nm^-2]",
            "Back pressure [Nm^-2]"
        }
        self.length_keys = {
            "Fuel length [m]",
            "Initial port diameter [m]",
            "Fuel outer diameter [m]",
            "Orifice diameter [m]",
            "Initial nozzle throat diameter [m]",
            "Nozzle exit diameter [m]"
        }
        self.volume_keys = {
            "Oxidizer filling volume [m^3]"
        }

        main_frame = ttk.Frame(root, padding=10)
        main_frame.pack(fill="both", expand=True)

        self.left_frame = ttk.Frame(main_frame, padding=5)
        self.left_frame.pack(side="left", fill="both", expand=True)

        self.right_frame = ttk.Frame(main_frame, padding=5)
        self.right_frame.pack(side="right", fill="both")
        self.log_text = scrolledtext.ScrolledText(self.right_frame, wrap=tk.WORD, width=50, height=30)
        self.log_text.pack(fill="both", expand=True)

        button_frame = ttk.Frame(root, padding=5)
        button_frame.pack(fill="x")
        ttk.Button(button_frame, text="初期値として保存", command=self.save_initial_values).pack(side="left", padx=5)

        run_button_frame = ttk.Frame(root, padding=5)
        run_button_frame.pack(fill="x")
        ttk.Button(run_button_frame, text="HYRES実行", command=self.run_hyres).pack(side="left", padx=5)

        initial_values = self.load_initial_values()
        self.display_initial_entries(initial_values)

        section_frame = ttk.LabelFrame(self.left_frame, text="output file name", padding=5)
        section_frame.pack(fill="x", pady=5)
        ttk.Label(section_frame, text="output excel file name", width=30).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.output_file_name_entry = ttk.Entry(section_frame)
        self.output_file_name_entry.grid(row=0, column=0 + 1, padx=5, pady=2, sticky="w")
        self.output_file_name_entry.insert(0, "result.xlsx")

    def log_message(self, message):
        self.log_text.insert("end", message+"\n")
        self.log_text.see("end")

    def save_json(self, file_name):
        try:
            data_to_save = {}
            for section, keys in self.entries.items():
                data_to_save[section] = {}
                for key, label in keys.items():
                    value = self.entry_data[(section, key)].get()
                    try:
                        if not value:
                            self.log_message(f"{label}の値が空欄です")
                            return -1
                        if key in self.pressure_keys:
                            value = float(value) * 1e6
                        elif key in self.length_keys:
                            value = float(value) / 1000
                        elif key in self.volume_keys:
                            value = float(value) / 1e6
                        else:
                            value = float(value)
                    except ValueError:
                        pass
                    data_to_save[section][key] = value
            with open(file_name, "w") as file:
                json.dump(data_to_save, file, indent=4)
            return 0
        except Exception as e:
            self.log_message(f"Error: {e}")
            return -1

    def save_initial_values(self):
        if self.save_json('initial_values.json') == 0:
            self.log_message("初期値を更新しました")
        else:
            self.log_message("初期値の更新に失敗しました")

    def load_initial_values(self):
        if os.path.exists("initial_values.json"):
            with open("initial_values.json", "r") as file:
                initial_values = json.load(file)
            return initial_values
        else:
            return {}

    def display_initial_entries(self, initial_values):
        for widget in self.left_frame.winfo_children():
            widget.destroy()

        for section_name, keys in self.entries.items():
            section_frame = ttk.LabelFrame(self.left_frame, text=section_name, padding=5)
            section_frame.pack(fill="x", pady=5)

            row = 0
            col = 0
            for key, label in keys.items():
                ttk.Label(section_frame, text=label, width=30).grid(row=row, column=col, padx=5, pady=2, sticky="w")
                entry = ttk.Entry(section_frame)
                entry.grid(row=row, column=col + 1, padx=5, pady=2, sticky="w")
                
                if section_name in initial_values and key in initial_values[section_name]:
                    value = initial_values[section_name][key]
                    if key in self.pressure_keys:
                        value = value / 1e6
                    elif key in self.length_keys:
                        value = value * 1000
                    elif key in self.volume_keys:
                        value = value * 1e6
                    entry.insert(0, value)

                self.entry_data[(section_name, key)] = entry

                if col == 0:
                    col = 2
                else:
                    col = 0
                    row += 1

    def run_hyres(self):
        def execute():
            try:
                process = subprocess.Popen(
                    "HYRES.exe",
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                for line in process.stdout:
                    self.log_message(line.strip())
                process.wait()
            except Exception as e:
                self.log_message(f"Exception: {e}")

        def check_thread_status():
            if self.thread and self.thread.is_alive():
                self.root.after(100, check_thread_status)
            else:
                post_process()

        def post_process():
            to_xlsx.make_output_xlsx(self.output_file_name_entry.get())
            if os.path.exists("input.json"):
                os.remove("input.json")
            if os.path.exists("output.csv"):
                os.remove("output.csv")
            self.log_message("燃焼計算が完了しました")

        self.log_message("燃焼計算を実行します")

        if self.save_json("input.json") == 0:
            self.thread = threading.Thread(target=execute)
            self.thread.daemon = True
            self.thread.start()
            check_thread_status()
        else:
            self.log_message("入力データの保存に失敗しました")
    
if __name__ == "__main__":
    root = tk.Tk()
    app = HYRESrunnerAPP(root)
    root.mainloop()
