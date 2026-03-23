"""
convert_teachers_local.py  –  Giao diện tkinter, KHÔNG cần API key
Yêu cầu: pip install openpyxl pandas
Chạy:    python convert_teachers_local.py
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from teacher_core import process_data

NIEN_KHOA_OPTIONS = ["2025-2026", "2026-2027", "2027-2028"]

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Chuyển đổi dữ liệu giáo viên")
        self.resizable(False, False)
        self._build_ui()

    def _build_ui(self):
        pad = {"padx":12,"pady":6}

        frm_in = tk.LabelFrame(self, text="File đầu vào", **pad)
        frm_in.pack(fill="x", **pad)
        self.var_input = tk.StringVar()
        tk.Entry(frm_in, textvariable=self.var_input, width=55,
                 state="readonly").pack(side="left", padx=(6,4), pady=6)
        tk.Button(frm_in, text="Chọn...", command=self._pick_input).pack(side="left")

        frm_out = tk.LabelFrame(self, text="File đầu ra", **pad)
        frm_out.pack(fill="x", **pad)
        self.var_output = tk.StringVar()
        tk.Entry(frm_out, textvariable=self.var_output, width=55,
                 state="readonly").pack(side="left", padx=(6,4), pady=6)
        tk.Button(frm_out, text="Lưu tại...", command=self._pick_output).pack(side="left")

        frm_nk = tk.LabelFrame(self, text="Niên khóa", **pad)
        frm_nk.pack(fill="x", **pad)
        self.var_nk = tk.StringVar(value=NIEN_KHOA_OPTIONS[0])
        ttk.Combobox(frm_nk, textvariable=self.var_nk,
                     values=NIEN_KHOA_OPTIONS, state="readonly",
                     width=14).pack(padx=6, pady=6, anchor="w")

        frm_log = tk.LabelFrame(self, text="Tiến trình", **pad)
        frm_log.pack(fill="both", expand=True, **pad)
        self.log_box = tk.Text(frm_log, height=10, width=70, state="disabled",
                               font=("Consolas",9))
        scroll = tk.Scrollbar(frm_log, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=scroll.set)
        self.log_box.pack(side="left", fill="both", expand=True, padx=(6,0), pady=6)
        scroll.pack(side="left", fill="y", pady=6)

        self.btn_run = tk.Button(self, text="▶  Chạy chuyển đổi",
                                 command=self._run, bg="#1F4E79", fg="white",
                                 font=("Arial",11,"bold"), pady=6)
        self.btn_run.pack(fill="x", padx=12, pady=(4,12))

    def _pick_input(self):
        p = filedialog.askopenfilename(
            title="Chọn file Excel đầu vào",
            filetypes=[("Excel files","*.xlsx *.xls *.xlsm"),("All files","*.*")])
        if p: self.var_input.set(p)

    def _pick_output(self):
        p = filedialog.asksaveasfilename(
            title="Lưu file đầu ra", defaultextension=".xlsx",
            filetypes=[("Excel files","*.xlsx"),("All files","*.*")])
        if p: self.var_output.set(p)

    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg+"\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _run(self):
        inp = self.var_input.get().strip()
        out = self.var_output.get().strip()
        nk  = self.var_nk.get()
        if not inp: messagebox.showerror("Lỗi","Vui lòng chọn file đầu vào!"); return
        if not out: messagebox.showerror("Lỗi","Vui lòng chọn nơi lưu file!"); return
        self.btn_run.configure(state="disabled", text="⏳  Đang xử lý...")
        threading.Thread(target=self._worker, args=(inp,out,nk), daemon=True).start()

    def _worker(self, inp, out, nk):
        try:
            result = process_data(inp, nk,
                                  progress_cb=lambda m: self.after(0, self._log, m))
            with open(out,"wb") as f: f.write(result)
            self.after(0, self._done_ok, out)
        except Exception as e:
            self.after(0, self._done_err, str(e))

    def _done_ok(self, out):
        self.btn_run.configure(state="normal", text="▶  Chạy chuyển đổi")
        messagebox.showinfo("Hoàn thành",f"✅ Chuyển đổi thành công!\nFile đã lưu:\n{out}")

    def _done_err(self, msg):
        self.btn_run.configure(state="normal", text="▶  Chạy chuyển đổi")
        self._log(f"❌ Lỗi: {msg}")
        messagebox.showerror("Lỗi", msg)

if __name__ == "__main__":
    App().mainloop()
