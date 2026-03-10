from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.services.pipeline import generate_general_report_from_vpn, generate_vpn_report


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("유지보수 보고서 작성 프로그램")
        self.geometry("980x560")
        self.minsize(920, 540)

        self.excel_path = tk.StringVar(value="")
        self.vpn_template_path = tk.StringVar(value="")
        self.edited_vpn_path = tk.StringVar(value="")
        self.general_template_path = tk.StringVar(value="")

        self._build_ui()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        ttk.Label(
            root,
            text="유지보수 보고서 작성 프로그램",
            font=("Malgun Gothic", 18, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))

        vpn_frame = ttk.LabelFrame(root, text="1단계. VPN 내역서 생성", padding=12)
        vpn_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")
        self._build_file_row(vpn_frame, 0, "유지보수 엑셀", self.excel_path, self._pick_excel)
        self._build_file_row(vpn_frame, 1, "VPN 템플릿", self.vpn_template_path, self._pick_vpn_template)
        ttk.Label(
            vpn_frame,
            text="엑셀 완료건을 기준으로 VPN 작업 내역서를 생성합니다.",
            foreground="#555",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Button(vpn_frame, text="VPN 내역서 생성", command=self._run_step1, width=18).grid(
            row=3, column=2, sticky="e", pady=(12, 0)
        )

        general_frame = ttk.LabelFrame(root, text="2단계. 일반 내역서 생성", padding=12)
        general_frame.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=(16, 0))
        self._build_file_row(general_frame, 0, "수정된 VPN 내역서", self.edited_vpn_path, self._pick_edited_vpn)
        self._build_file_row(
            general_frame, 1, "일반 템플릿", self.general_template_path, self._pick_general_template
        )
        ttk.Label(
            general_frame,
            text="수정된 VPN 내역서의 작업내역 2줄 구조를 읽어 일반 유지보수 내역서로 변환합니다.",
            foreground="#555",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Button(general_frame, text="일반 내역서 생성", command=self._run_step2, width=18).grid(
            row=3, column=2, sticky="e", pady=(12, 0)
        )

        tips = (
            "작업 순서:\n"
            "1) 엑셀과 VPN 템플릿으로 VPN 내역서를 생성합니다.\n"
            "2) Word에서 VPN 작업내역을 수정합니다.\n"
            "3) 수정된 VPN 내역서와 일반 템플릿으로 일반 유지보수 내역서를 생성합니다."
        )
        ttk.Label(root, text=tips, foreground="#444").grid(row=3, column=0, columnspan=3, sticky="w", pady=(16, 0))

        root.columnconfigure(0, weight=1)
        vpn_frame.columnconfigure(1, weight=1)
        general_frame.columnconfigure(1, weight=1)

    def _build_file_row(self, parent, row: int, label: str, variable: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label, font=("Malgun Gothic", 11)).grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(parent, text="선택", command=command, width=12).grid(row=row, column=2, sticky="e")

    def _pick_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="유지보수 엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.excel_path.set(path)

    def _pick_vpn_template(self) -> None:
        path = filedialog.askopenfilename(
            title="VPN 템플릿 파일 선택",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.vpn_template_path.set(path)

    def _pick_edited_vpn(self) -> None:
        path = filedialog.askopenfilename(
            title="수정된 VPN 내역서 선택",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.edited_vpn_path.set(path)

    def _pick_general_template(self) -> None:
        path = filedialog.askopenfilename(
            title="일반 유지보수 템플릿 선택",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.general_template_path.set(path)

    def _require_existing_path(self, path_text: str, label: str) -> Path | None:
        path = Path(path_text.strip())
        if not path_text.strip() or not path.exists():
            messagebox.showerror("오류", f"유효한 {label} 파일을 선택해 주세요.")
            return None
        return path

    def _run_step1(self) -> None:
        excel = self._require_existing_path(self.excel_path.get(), "유지보수 엑셀")
        vpn_template = self._require_existing_path(self.vpn_template_path.get(), "VPN 템플릿")
        if excel is None or vpn_template is None:
            return

        try:
            result = generate_vpn_report(excel, vpn_template)
        except Exception as exc:
            messagebox.showerror("실패", f"VPN 내역서 생성 중 오류가 발생했습니다.\n\n{exc}")
            return

        self.edited_vpn_path.set(str(result.output_path))
        messagebox.showinfo(
            "완료",
            f"VPN 내역서 생성 완료\n\n저장 위치:\n{result.output_path}\n\n검증 통과: {result.validated_count}건",
        )

    def _run_step2(self) -> None:
        edited_vpn = self._require_existing_path(self.edited_vpn_path.get(), "수정된 VPN 내역서")
        general_template = self._require_existing_path(self.general_template_path.get(), "일반 템플릿")
        if edited_vpn is None or general_template is None:
            return

        try:
            result = generate_general_report_from_vpn(edited_vpn, general_template)
        except Exception as exc:
            messagebox.showerror("실패", f"일반 내역서 생성 중 오류가 발생했습니다.\n\n{exc}")
            return

        message = (
            f"일반 유지보수 내역서 생성 완료\n\n저장 위치:\n{result.output_path}\n\n검증 통과: {result.validated_count}건"
        )
        if result.warnings:
            preview = "\n".join(result.warnings[:5])
            if len(result.warnings) > 5:
                preview += f"\n... 외 {len(result.warnings) - 5}건"
            message += f"\n\n파싱 경고:\n{preview}"

        messagebox.showinfo("완료", message)


def run() -> None:
    App().mainloop()
