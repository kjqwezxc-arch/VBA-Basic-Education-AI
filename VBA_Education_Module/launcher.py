#!/usr/bin/env python3
"""
VBA Macro Education Module - Launcher
Yamaha Motor Co., Ltd. | DX Education 2026
Author: DX Education Team
"""

import tkinter as tk
from tkinter import ttk, messagebox, font
import webbrowser
import os
import json
import sys
import datetime
from pathlib import Path

# ─────────────────────────────────────────────
#  MODULE DEFINITIONS
# ─────────────────────────────────────────────
MODULES = [
    {
        "id": "01", "num": 1,
        "title": "Introduction to VBA",
        "subtitle": "What is VBA and why learn it?",
        "duration": "20 min", "difficulty": "Beginner",
        "icon": "📘", "color": "#1565c0",
        "file": "modules/01_intro.html",
        "topics": ["What is VBA?", "Why automate?", "Real-world examples", "Course overview"],
    },
    {
        "id": "02", "num": 2,
        "title": "Getting Started with VBE",
        "subtitle": "Developer Tab, VB Editor, First Macro",
        "duration": "30 min", "difficulty": "Beginner",
        "icon": "🚀", "color": "#0277bd",
        "file": "modules/02_getting_started.html",
        "topics": ["Enable Developer Tab", "VBE Interface", "Record a Macro", "Your first Sub"],
    },
    {
        "id": "03", "num": 3,
        "title": "Variables & Data Types",
        "subtitle": "Dim, data types, Option Explicit",
        "duration": "35 min", "difficulty": "Beginner",
        "icon": "📦", "color": "#00695c",
        "file": "modules/03_variables.html",
        "topics": ["Dim statement", "Data types", "Option Explicit", "Constants"],
    },
    {
        "id": "04", "num": 4,
        "title": "Control Flow",
        "subtitle": "If/Then, Loops, Select Case",
        "duration": "40 min", "difficulty": "Beginner",
        "icon": "🔀", "color": "#4527a0",
        "file": "modules/04_control_flow.html",
        "topics": ["If/Then/Else", "For/Next", "For Each", "Do Loops", "Select Case"],
    },
    {
        "id": "05", "num": 5,
        "title": "Procedures & Functions",
        "subtitle": "Sub, Function, parameters, scope",
        "duration": "35 min", "difficulty": "Intermediate",
        "icon": "⚙️", "color": "#558b2f",
        "file": "modules/05_procedures.html",
        "topics": ["Sub vs Function", "Parameters", "ByVal vs ByRef", "Scope"],
    },
    {
        "id": "06", "num": 6,
        "title": "Excel Object Model",
        "subtitle": "Workbooks, Worksheets, Ranges, Cells",
        "duration": "45 min", "difficulty": "Intermediate",
        "icon": "📊", "color": "#e65100",
        "file": "modules/06_excel_objects.html",
        "topics": ["Object hierarchy", "Workbooks", "Worksheets", "Range & Cells"],
    },
    {
        "id": "07", "num": 7,
        "title": "Built-in Functions",
        "subtitle": "String, Math, Date, MsgBox, InputBox",
        "duration": "30 min", "difficulty": "Intermediate",
        "icon": "🔧", "color": "#5d4037",
        "file": "modules/07_functions.html",
        "topics": ["String functions", "Math functions", "Date functions", "MsgBox/InputBox"],
    },
    {
        "id": "08", "num": 8,
        "title": "Error Handling & Debugging",
        "subtitle": "On Error, Debugging tools, Best practices",
        "duration": "30 min", "difficulty": "Intermediate",
        "icon": "🐛", "color": "#c62828",
        "file": "modules/08_error_handling.html",
        "topics": ["Error types", "On Error GoTo", "Debug tools", "Immediate Window"],
    },
    {
        "id": "09", "num": 9,
        "title": "UserForms & UI",
        "subtitle": "Creating forms, controls, events",
        "duration": "40 min", "difficulty": "Intermediate",
        "icon": "🖥️", "color": "#1a237e",
        "file": "modules/09_userforms.html",
        "topics": ["Creating UserForms", "Controls", "Events", "Validation"],
    },
    {
        "id": "10", "num": 10,
        "title": "Practical Applications",
        "subtitle": "Real-world macros and examples",
        "duration": "45 min", "difficulty": "Intermediate",
        "icon": "💼", "color": "#33691e",
        "file": "modules/10_practical.html",
        "topics": ["Data processing", "Report generation", "Automation", "Best practices"],
    },
    {
        "id": "11", "num": 11,
        "title": "AI for VBA – Introduction",
        "subtitle": "ChatGPT, Copilot, writing prompts",
        "duration": "35 min", "difficulty": "Advanced",
        "icon": "🤖", "color": "#880e4f",
        "file": "modules/11_ai_intro.html",
        "topics": ["AI tools overview", "ChatGPT for VBA", "Writing prompts", "Testing AI code"],
    },
    {
        "id": "12", "num": 12,
        "title": "AI for VBA – Advanced",
        "subtitle": "Prompt engineering, debugging with AI",
        "duration": "40 min", "difficulty": "Advanced",
        "icon": "🧠", "color": "#4a148c",
        "file": "modules/12_ai_advanced.html",
        "topics": ["Prompt engineering", "AI debugging", "Code review", "Full projects"],
    },
]

DIFFICULTY_COLOR = {"Beginner": "#4caf50", "Intermediate": "#ff9800", "Advanced": "#f44336"}

BASE_DIR = Path(__file__).parent
PROGRESS_FILE = BASE_DIR / "progress.json"


# ─────────────────────────────────────────────
#  MAIN APP CLASS
# ─────────────────────────────────────────────
class VBALauncher:
    def __init__(self):
        self.progress = self.load_progress()

        self.root = tk.Tk()
        self.root.title("VBA Macro Education – Learning Portal")
        self.root.geometry("1100x720")
        self.root.minsize(900, 600)
        self.root.configure(bg="#0d1117")

        # Icon
        try:
            self.root.iconbitmap(default="")
        except Exception:
            pass

        self.setup_styles()
        self.build_ui()
        self.refresh_cards()
        self.root.mainloop()

    # ── Progress I/O ──────────────────────────
    def load_progress(self):
        if PROGRESS_FILE.exists():
            try:
                with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"completed": [], "accessed": {}, "notes": {}, "started_date": str(datetime.date.today())}

    def save_progress(self):
        try:
            with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.progress, f, indent=2, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save progress:\n{e}")

    # ── Styles ────────────────────────────────
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook", background="#0d1117", borderwidth=0)
        style.configure("TNotebook.Tab", background="#161b22", foreground="#8b949e",
                        padding=[16, 8], font=("Segoe UI", 10))
        style.map("TNotebook.Tab", background=[("selected", "#1f2937")],
                  foreground=[("selected", "#ffffff")])
        style.configure("Vertical.TScrollbar", background="#1f2937", troughcolor="#0d1117",
                        arrowcolor="#8b949e")

    # ── UI Builder ────────────────────────────
    def build_ui(self):
        # ─ Header ─
        hdr = tk.Frame(self.root, bg="#161b22", height=70)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text="VBA Macro", font=("Segoe UI", 18, "bold"),
                 bg="#161b22", fg="#ff8f00").pack(side="left", padx=(20, 0), pady=15)
        tk.Label(hdr, text=" Education Module",
                 font=("Segoe UI", 18), bg="#161b22", fg="#e6edf3").pack(side="left")

        tk.Label(hdr, text="DX Education 2026  |  Yamaha Motor Co., Ltd.",
                 font=("Segoe UI", 9), bg="#161b22", fg="#8b949e").pack(side="right", padx=20)

        # ─ Thin accent line ─
        tk.Frame(self.root, bg="#ff8f00", height=2).pack(fill="x")

        # ─ Notebook ─
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=0, pady=0)

        self.tab_home = tk.Frame(self.nb, bg="#0d1117")
        self.tab_modules = tk.Frame(self.nb, bg="#0d1117")
        self.tab_progress = tk.Frame(self.nb, bg="#0d1117")
        self.tab_notes = tk.Frame(self.nb, bg="#0d1117")

        self.nb.add(self.tab_home, text="  🏠 Home  ")
        self.nb.add(self.tab_modules, text="  📚 Modules  ")
        self.nb.add(self.tab_progress, text="  📈 Progress  ")
        self.nb.add(self.tab_notes, text="  📝 My Notes  ")

        self.build_home_tab()
        self.build_modules_tab()
        self.build_progress_tab()
        self.build_notes_tab()

        # ─ Status bar ─
        self.status_var = tk.StringVar(value="Ready  |  Click a module to start learning")
        sb = tk.Frame(self.root, bg="#161b22", height=28)
        sb.pack(fill="x", side="bottom")
        tk.Label(sb, textvariable=self.status_var, bg="#161b22", fg="#8b949e",
                 font=("Segoe UI", 9), anchor="w").pack(side="left", padx=12)
        completed = len(self.progress.get("completed", []))
        self.progress_label = tk.Label(sb, text=f"Progress: {completed}/12 modules",
                                       bg="#161b22", fg="#ff8f00",
                                       font=("Segoe UI", 9, "bold"))
        self.progress_label.pack(side="right", padx=12)

    # ── Home Tab ──────────────────────────────
    def build_home_tab(self):
        frame = self.tab_home
        tk.Label(frame, text="\n👋  Welcome to VBA Macro Education!",
                 font=("Segoe UI", 20, "bold"), bg="#0d1117", fg="#e6edf3").pack(pady=(30, 0))
        tk.Label(frame,
                 text="A complete, self-paced learning path to master Excel VBA — from zero to AI-assisted coding.",
                 font=("Segoe UI", 11), bg="#0d1117", fg="#8b949e", wraplength=700).pack(pady=(8, 24))

        # Stats row
        stats = tk.Frame(frame, bg="#0d1117")
        stats.pack(pady=10)
        self._stat_box(stats, "📚", "12", "Modules").pack(side="left", padx=12)
        self._stat_box(stats, "⏱️", "6.5h", "Total Time").pack(side="left", padx=12)
        self._stat_box(stats, "🎯", "60+", "Code Examples").pack(side="left", padx=12)
        self._stat_box(stats, "🤖", "2", "AI Modules").pack(side="left", padx=12)

        # Progress bar
        completed = len(self.progress.get("completed", []))
        pct = int(completed / 12 * 100)
        prog_frame = tk.Frame(frame, bg="#0d1117")
        prog_frame.pack(pady=20, padx=60, fill="x")
        tk.Label(prog_frame, text=f"Your Progress:  {completed}/12 modules ({pct}%)",
                 font=("Segoe UI", 11), bg="#0d1117", fg="#e6edf3").pack(anchor="w")
        bar_bg = tk.Frame(prog_frame, bg="#21262d", height=14, relief="flat")
        bar_bg.pack(fill="x", pady=4)
        if pct > 0:
            bar_fill = tk.Frame(bar_bg, bg="#ff8f00", height=14)
            bar_fill.place(relwidth=pct / 100, relheight=1)

        # Learning path
        tk.Label(frame, text="📋  Recommended Learning Path",
                 font=("Segoe UI", 13, "bold"), bg="#0d1117", fg="#ff8f00").pack(pady=(20, 8))
        path_txt = ("Modules 1–4  →  VBA Fundamentals (Beginner)\n"
                    "Modules 5–8  →  Core Programming Skills (Intermediate)\n"
                    "Modules 9–10  →  Real-World Applications\n"
                    "Modules 11–12  →  AI-Assisted VBA Development (Advanced)")
        tk.Label(frame, text=path_txt, font=("Segoe UI", 10), bg="#0d1117", fg="#8b949e",
                 justify="left").pack(padx=60)

        # Quick start button
        tk.Button(frame, text="  🚀  Start Learning from Module 1  ",
                  font=("Segoe UI", 12, "bold"), bg="#ff8f00", fg="#000",
                  relief="flat", padx=20, pady=10, cursor="hand2",
                  command=lambda: self.open_module(MODULES[0])).pack(pady=30)

    def _stat_box(self, parent, icon, value, label):
        box = tk.Frame(parent, bg="#161b22", padx=20, pady=14)
        tk.Label(box, text=icon, font=("Segoe UI", 22), bg="#161b22").pack()
        tk.Label(box, text=value, font=("Segoe UI", 18, "bold"), bg="#161b22", fg="#ff8f00").pack()
        tk.Label(box, text=label, font=("Segoe UI", 9), bg="#161b22", fg="#8b949e").pack()
        return box

    # ── Modules Tab ───────────────────────────
    def build_modules_tab(self):
        # Scrollable canvas
        canvas = tk.Canvas(self.tab_modules, bg="#0d1117", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.tab_modules, orient="vertical", command=canvas.yview)
        self.modules_frame = tk.Frame(canvas, bg="#0d1117")

        self.modules_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.modules_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    def refresh_cards(self):
        for w in self.modules_frame.winfo_children():
            w.destroy()

        tk.Label(self.modules_frame, text="  📚  All Modules",
                 font=("Segoe UI", 14, "bold"), bg="#0d1117", fg="#e6edf3").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=20, pady=(16, 8))

        for i, mod in enumerate(MODULES):
            row = (i // 3) + 1
            col = i % 3
            self._make_card(self.modules_frame, mod, row, col)

        # Padding row
        tk.Label(self.modules_frame, text="", bg="#0d1117").grid(row=100, column=0, pady=20)

    def _make_card(self, parent, mod, row, col):
        completed = mod["id"] in self.progress.get("completed", [])
        accessed = mod["id"] in self.progress.get("accessed", {})

        card = tk.Frame(parent, bg="#161b22", padx=16, pady=14, cursor="hand2",
                        relief="flat", bd=0)
        card.grid(row=row, column=col, padx=12, pady=8, sticky="nsew")
        parent.columnconfigure(col, weight=1)

        # Colored top strip
        strip = tk.Frame(card, bg=mod["color"], height=4)
        strip.pack(fill="x", pady=(0, 10))

        # Header row
        hrow = tk.Frame(card, bg="#161b22")
        hrow.pack(fill="x")
        tk.Label(hrow, text=f"Module {mod['num']:02d}", font=("Segoe UI", 8, "bold"),
                 bg="#161b22", fg="#8b949e").pack(side="left")
        status_txt = "✅" if completed else ("📖" if accessed else "⬜")
        tk.Label(hrow, text=status_txt, bg="#161b22", font=("Segoe UI", 11)).pack(side="right")

        tk.Label(card, text=f"{mod['icon']}  {mod['title']}",
                 font=("Segoe UI", 11, "bold"), bg="#161b22", fg="#e6edf3",
                 wraplength=250, justify="left").pack(anchor="w", pady=(2, 4))
        tk.Label(card, text=mod["subtitle"], font=("Segoe UI", 9),
                 bg="#161b22", fg="#8b949e", wraplength=250, justify="left").pack(anchor="w")

        # Tags row
        trow = tk.Frame(card, bg="#161b22")
        trow.pack(fill="x", pady=(8, 4))
        diff_color = DIFFICULTY_COLOR.get(mod["difficulty"], "#888")
        tk.Label(trow, text=f" {mod['difficulty']} ",
                 bg=diff_color, fg="#000", font=("Segoe UI", 8, "bold"),
                 padx=4).pack(side="left")
        tk.Label(trow, text=f"  ⏱ {mod['duration']}",
                 font=("Segoe UI", 9), bg="#161b22", fg="#8b949e").pack(side="left")

        # Open button
        btn = tk.Button(card, text="▶  Open Module",
                        font=("Segoe UI", 9, "bold"),
                        bg=mod["color"], fg="#fff",
                        relief="flat", padx=10, pady=4, cursor="hand2",
                        command=lambda m=mod: self.open_module(m))
        btn.pack(fill="x", pady=(10, 2))

        if completed:
            tk.Button(card, text="↩  Mark Incomplete",
                      font=("Segoe UI", 8), bg="#21262d", fg="#8b949e",
                      relief="flat", padx=6, pady=2, cursor="hand2",
                      command=lambda m=mod: self.toggle_complete(m)).pack(fill="x")
        else:
            tk.Button(card, text="✓  Mark as Complete",
                      font=("Segoe UI", 8), bg="#1b4332", fg="#4caf50",
                      relief="flat", padx=6, pady=2, cursor="hand2",
                      command=lambda m=mod: self.toggle_complete(m)).pack(fill="x")

        # Hover effect
        for widget in card.winfo_children():
            widget.bind("<Enter>", lambda e, c=card: c.configure(bg="#1f2937"))
            widget.bind("<Leave>", lambda e, c=card: c.configure(bg="#161b22"))

    def open_module(self, mod):
        module_path = BASE_DIR / mod["file"]
        if not module_path.exists():
            messagebox.showerror("File Not Found",
                                 f"Module file not found:\n{module_path}\n\n"
                                 "Please ensure all module files are present.")
            return

        # Track access
        accessed = self.progress.setdefault("accessed", {})
        accessed[mod["id"]] = str(datetime.datetime.now())
        self.save_progress()
        self.refresh_cards()

        # Open in browser
        webbrowser.open(module_path.as_uri())
        self.status_var.set(f"Opened: Module {mod['num']} – {mod['title']}")
        self.nb.select(self.tab_modules)

    def toggle_complete(self, mod):
        completed = self.progress.setdefault("completed", [])
        if mod["id"] in completed:
            completed.remove(mod["id"])
            self.status_var.set(f"Module {mod['num']} marked as incomplete")
        else:
            if mod["id"] not in completed:
                completed.append(mod["id"])
            self.status_var.set(f"✅ Module {mod['num']} marked as complete!")
        self.save_progress()
        self.refresh_cards()
        self.refresh_progress_tab()
        c = len(completed)
        self.progress_label.configure(text=f"Progress: {c}/12 modules")

    # ── Progress Tab ──────────────────────────
    def build_progress_tab(self):
        self.prog_tab_frame = tk.Frame(self.tab_progress, bg="#0d1117")
        self.prog_tab_frame.pack(fill="both", expand=True, padx=20, pady=20)
        self.refresh_progress_tab()

    def refresh_progress_tab(self):
        for w in self.prog_tab_frame.winfo_children():
            w.destroy()

        completed = self.progress.get("completed", [])
        accessed = self.progress.get("accessed", {})
        n_done = len(completed)
        pct = int(n_done / 12 * 100)

        tk.Label(self.prog_tab_frame, text="📈  Your Progress",
                 font=("Segoe UI", 16, "bold"), bg="#0d1117", fg="#e6edf3").pack(anchor="w")
        tk.Label(self.prog_tab_frame, text=f"Started: {self.progress.get('started_date', 'N/A')}",
                 font=("Segoe UI", 10), bg="#0d1117", fg="#8b949e").pack(anchor="w", pady=(4, 16))

        # Big number
        tk.Label(self.prog_tab_frame, text=f"{n_done} / 12",
                 font=("Segoe UI", 48, "bold"), bg="#0d1117", fg="#ff8f00").pack()
        tk.Label(self.prog_tab_frame, text=f"modules completed  ({pct}%)",
                 font=("Segoe UI", 12), bg="#0d1117", fg="#8b949e").pack(pady=(0, 16))

        # Progress bar
        bar_bg = tk.Frame(self.prog_tab_frame, bg="#21262d", height=20)
        bar_bg.pack(fill="x", padx=40, pady=4)
        if pct > 0:
            bar_fill = tk.Frame(bar_bg, bg="#ff8f00", height=20)
            bar_fill.place(relwidth=pct / 100, relheight=1)

        # Module checklist
        tk.Label(self.prog_tab_frame, text="\n  Module Checklist",
                 font=("Segoe UI", 12, "bold"), bg="#0d1117", fg="#e6edf3").pack(anchor="w")

        for mod in MODULES:
            done = mod["id"] in completed
            seen = mod["id"] in accessed
            row = tk.Frame(self.prog_tab_frame, bg="#0d1117")
            row.pack(fill="x", padx=10, pady=2)

            status = "✅" if done else ("👁️" if seen else "⬜")
            color = "#4caf50" if done else ("#ff9800" if seen else "#444")
            tk.Label(row, text=f"  {status}  Module {mod['num']:02d} – {mod['title']}",
                     font=("Segoe UI", 10), bg="#0d1117", fg=color, anchor="w").pack(side="left")

        # Reset button
        tk.Button(self.prog_tab_frame, text="🔄  Reset All Progress",
                  font=("Segoe UI", 10), bg="#3d0000", fg="#f44336",
                  relief="flat", padx=12, pady=6, cursor="hand2",
                  command=self.reset_progress).pack(pady=30)

    def reset_progress(self):
        if messagebox.askyesno("Reset Progress",
                               "Are you sure you want to reset all progress?\nThis cannot be undone."):
            self.progress = {"completed": [], "accessed": {}, "notes": {},
                             "started_date": str(datetime.date.today())}
            self.save_progress()
            self.refresh_cards()
            self.refresh_progress_tab()
            self.status_var.set("Progress reset.")
            self.progress_label.configure(text="Progress: 0/12 modules")

    # ── Notes Tab ─────────────────────────────
    def build_notes_tab(self):
        frame = self.tab_notes
        tk.Label(frame, text="  📝  My Notes",
                 font=("Segoe UI", 14, "bold"), bg="#0d1117", fg="#e6edf3").pack(
            anchor="w", padx=20, pady=(16, 4))
        tk.Label(frame, text="  Jot down key learnings from each module.",
                 font=("Segoe UI", 10), bg="#0d1117", fg="#8b949e").pack(anchor="w", padx=20)

        # Module selector
        sel_frame = tk.Frame(frame, bg="#0d1117")
        sel_frame.pack(fill="x", padx=20, pady=(12, 4))
        tk.Label(sel_frame, text="Module:", font=("Segoe UI", 10), bg="#0d1117", fg="#e6edf3").pack(side="left")
        self.note_mod_var = tk.StringVar(value=MODULES[0]["title"])
        mod_names = [f"Module {m['num']:02d}: {m['title']}" for m in MODULES]
        self.note_dropdown = ttk.Combobox(sel_frame, values=mod_names, width=40,
                                          state="readonly")
        self.note_dropdown.current(0)
        self.note_dropdown.pack(side="left", padx=8)
        self.note_dropdown.bind("<<ComboboxSelected>>", self.load_note)

        # Text area
        txt_frame = tk.Frame(frame, bg="#0d1117")
        txt_frame.pack(fill="both", expand=True, padx=20, pady=8)
        self.note_text = tk.Text(txt_frame, bg="#161b22", fg="#e6edf3",
                                 font=("Segoe UI", 10), insertbackground="#ff8f00",
                                 relief="flat", padx=12, pady=10, wrap="word",
                                 undo=True)
        note_scroll = ttk.Scrollbar(txt_frame, orient="vertical",
                                    command=self.note_text.yview)
        self.note_text.configure(yscrollcommand=note_scroll.set)
        self.note_text.pack(side="left", fill="both", expand=True)
        note_scroll.pack(side="right", fill="y")

        # Save button
        tk.Button(frame, text="💾  Save Notes",
                  font=("Segoe UI", 10, "bold"), bg="#1b4332", fg="#4caf50",
                  relief="flat", padx=16, pady=6, cursor="hand2",
                  command=self.save_note).pack(pady=8)

        self.load_note()

    def _get_note_key(self):
        idx = self.note_dropdown.current()
        return MODULES[idx]["id"] if idx >= 0 else "00"

    def load_note(self, event=None):
        key = self._get_note_key()
        text = self.progress.get("notes", {}).get(key, "")
        self.note_text.delete("1.0", "end")
        self.note_text.insert("1.0", text)

    def save_note(self):
        key = self._get_note_key()
        notes = self.progress.setdefault("notes", {})
        notes[key] = self.note_text.get("1.0", "end-1c")
        self.save_progress()
        self.status_var.set("Notes saved ✓")


# ─────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    try:
        app = VBALauncher()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
