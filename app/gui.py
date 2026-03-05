from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.services.pipeline import generate_report


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("\uc720\uc9c0\ubcf4\uc218 \ubcf4\uace0\uc11c \uc791\uc131 \ud504\ub85c\uadf8\ub7a8")
        self.geometry("900x320")
        self.minsize(820, 300)

        self.excel_path = tk.StringVar(value="")
        self.word_path = tk.StringVar(value="")

        self._build_ui()

    def _build_ui(self) -> None:
        pad = 10
        root = ttk.Frame(self, padding=pad)
        root.pack(fill="both", expand=True)

        title = ttk.Label(
            root,
            text="\uc720\uc9c0\ubcf4\uc218 \ubcf4\uace0\uc11c \uc791\uc131 \ud504\ub85c\uadf8\ub7a8",
            font=("Malgun Gothic", 18, "bold"),
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 18))

        ttk.Label(root, text="\uc720\uc9c0\ubcf4\uc218 \ub0b4\uc5ed", font=("Malgun Gothic", 12)).grid(
            row=1, column=0, sticky="w", pady=8
        )
        excel_entry = ttk.Entry(root, textvariable=self.excel_path)
        excel_entry.grid(row=1, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(root, text="\uc5d1\uc140", command=self._pick_excel, width=12).grid(row=1, column=2, sticky="e")

        ttk.Label(root, text="\uc6cc\ub4dc \ud15c\ud50c\ub9bf", font=("Malgun Gothic", 12)).grid(
            row=2, column=0, sticky="w", pady=8
        )
        word_entry = ttk.Entry(root, textvariable=self.word_path)
        word_entry.grid(row=2, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(root, text="\uc6cc\ub4dc", command=self._pick_word, width=12).grid(row=2, column=2, sticky="e")

        ttk.Separator(root).grid(row=3, column=0, columnspan=3, sticky="ew", pady=18)
        ttk.Button(root, text="\uc791\uc131", command=self._run, width=16).grid(row=4, column=2, sticky="e")

        tips = (
            "\uc5d1\uc140\uacfc \uc6cc\ub4dc \ud30c\uc77c\uc744 \uc120\ud0dd\ud55c \ub4a4 '\uc791\uc131'\uc744 \ub204\ub974\uba74 \ubcf4\uace0\uc11c\ub97c \uc791\uc131\ud569\ub2c8\ub2e4.\n"
            "\uc9c4\ud589\uc0c1\ud0dc\uac00 '\uc644\ub8cc'\uc778 \ub370\uc774\ud130\ub9cc \ub9e4\ud551\ud558\uba70, \uc791\uc131 \uc644\ub8cc \ud6c4 \uc5d1\uc140-\uc6cc\ub4dc \uac04 \uac74\uc218/\ub0b4\uc6a9 \uad50\ucc28\uac80\uc99d\uc744 \uc790\ub3d9 \uc218\ud589\ud569\ub2c8\ub2e4."
        )
        ttk.Label(root, text=tips, foreground="#555").grid(row=5, column=0, columnspan=3, sticky="w", pady=(18, 0))
        root.columnconfigure(1, weight=1)

    def _pick_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="\uc720\uc9c0\ubcf4\uc218 \uc5d1\uc140 \ud30c\uc77c \uc120\ud0dd",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.excel_path.set(path)

    def _pick_word(self) -> None:
        path = filedialog.askopenfilename(
            title="\uc6cc\ub4dc \ud15c\ud50c\ub9bf \ud30c\uc77c \uc120\ud0dd",
            filetypes=[("Word files", "*.docx *.doc"), ("All files", "*.*")],
        )
        if path:
            self.word_path.set(path)

    def _run(self) -> None:
        excel = self.excel_path.get().strip()
        word = self.word_path.get().strip()

        if not excel or not Path(excel).exists():
            messagebox.showerror("\uc624\ub958", "\uc720\ud6a8\ud55c \uc5d1\uc140 \ud30c\uc77c\uc744 \uc120\ud0dd\ud574\uc8fc\uc138\uc694.")
            return
        if not word or not Path(word).exists():
            messagebox.showerror("\uc624\ub958", "\uc720\ud6a8\ud55c \uc6cc\ub4dc \ud15c\ud50c\ub9bf \ud30c\uc77c\uc744 \uc120\ud0dd\ud574\uc8fc\uc138\uc694.")
            return

        try:
            result = generate_report(Path(excel), Path(word))
        except Exception as e:
            messagebox.showerror("\uc2e4\ud328", f"\ubcf4\uace0\uc11c \uc791\uc131 \uc911 \uc624\ub958\uac00 \ubc1c\uc0dd\ud588\uc2b5\ub2c8\ub2e4.\n\n{e}")
            return

        messagebox.showinfo(
            "\uc644\ub8cc",
            f"\ubcf4\uace0\uc11c \uc791\uc131 \uc644\ub8cc!\n\n"
            f"\uc800\uc7a5 \uc704\uce58:\n{result.output_path}\n\n"
            f"\uad50\ucc28\uac80\uc99d \ud1b5\uacfc: {result.validated_count}\uac74",
        )


def run() -> None:
    App().mainloop()
