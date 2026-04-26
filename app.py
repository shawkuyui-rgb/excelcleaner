from __future__ import annotations

import os
import subprocess
import sys
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from excel_cleaner import CleanupRuleError, clean_workbook


def _default_ui_font() -> str:
    if os.name == "nt":
        return "Segoe UI"
    if sys.platform == "darwin":
        return "Avenir Next"
    return "DejaVu Sans"


def _default_mono_font() -> str:
    if os.name == "nt":
        return "Consolas"
    if sys.platform == "darwin":
        return "Menlo"
    return "DejaVu Sans Mono"


class ExcelCleanerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel 清理工具")
        self.root.geometry("860x620")
        self.root.minsize(760, 560)

        self.selected_file = tk.StringVar()
        self.file_name = tk.StringVar(value="尚未选择文件")
        self.output_dir_var = tk.StringVar(value="选择文件后自动生成")
        self.status_title = tk.StringVar(value="待开始")
        self.status_body = tk.StringVar(value="先选择一个 Excel 文件。")

        self.run_status = "idle"
        self.output_dir: Path | None = None
        self.last_output_file: Path | None = None

        self._build_styles()
        self._build_ui()
        self._update_run_button_state()
        self._refresh_status_card()

    def _build_styles(self) -> None:
        self.root.configure(bg="#eef4f7")
        style = ttk.Style()
        style.theme_use("clam")
        ui_font = _default_ui_font()
        mono_font = _default_mono_font()

        style.configure("App.TFrame", background="#eef4f7")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Info.TFrame", background="#f5f9fb")
        style.configure("Idle.TFrame", background="#f4f8fa")
        style.configure("Running.TFrame", background="#e8f1f4")
        style.configure("Success.TFrame", background="#e6f5f1")
        style.configure("Error.TFrame", background="#fdeceb")

        style.configure("Title.TLabel", font=(ui_font, 28, "bold"), background="#eef4f7", foreground="#102631")
        style.configure("Subtitle.TLabel", font=(ui_font, 12), background="#eef4f7", foreground="#607784")
        style.configure("CardTitle.TLabel", font=(ui_font, 14, "bold"), background="#ffffff", foreground="#163543")
        style.configure("CardBody.TLabel", font=(ui_font, 11), background="#ffffff", foreground="#5c7380")
        style.configure("InfoLabel.TLabel", font=(ui_font, 10, "bold"), background="#f5f9fb", foreground="#7b8f9b")
        style.configure("InfoValue.TLabel", font=(ui_font, 15, "bold"), background="#f5f9fb", foreground="#163543")
        style.configure("PathValue.TLabel", font=(mono_font, 10), background="#f5f9fb", foreground="#183545")
        style.configure("StatusTitle.TLabel", font=(ui_font, 18, "bold"), background="#f4f8fa", foreground="#163543")
        style.configure("StatusBody.TLabel", font=(ui_font, 11), background="#f4f8fa", foreground="#607784")
        style.configure("Primary.TButton", font=(ui_font, 12, "bold"), padding=(20, 12))
        style.configure("Secondary.TButton", font=(ui_font, 11, "bold"), padding=(16, 12))
        style.configure("Ghost.TButton", font=(ui_font, 11, "bold"), padding=(16, 12))

        style.map(
            "Primary.TButton",
            background=[("disabled", "#dbe6ea"), ("active", "#0c4b58"), ("!disabled", "#0e5a67")],
            foreground=[("disabled", "#95a8b2"), ("!disabled", "#ffffff")],
        )
        style.map(
            "Secondary.TButton",
            background=[("disabled", "#e8eef1"), ("active", "#d8e5ea"), ("!disabled", "#e4edf1")],
            foreground=[("disabled", "#97aab4"), ("!disabled", "#14313d")],
        )
        style.map(
            "Ghost.TButton",
            background=[("disabled", "#e8eef1"), ("active", "#f2f6f8"), ("!disabled", "#ffffff")],
            foreground=[("disabled", "#97aab4"), ("!disabled", "#5b7380")],
        )

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, style="App.TFrame", padding=28)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)

        header = ttk.Frame(container, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Excel 清理工具", style="Title.TLabel").pack(anchor="center")
        ttk.Label(
            header,
            text="选择文件，开始清理，打开结果目录。",
            style="Subtitle.TLabel",
            justify="center",
        ).pack(anchor="center", pady=(8, 0))

        card = ttk.Frame(container, style="Card.TFrame", padding=28)
        card.grid(row=1, column=0, sticky="nsew", pady=(24, 0))
        card.columnconfigure(0, weight=1)

        ttk.Label(card, text="当前文件", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        file_box = ttk.Frame(card, style="Info.TFrame", padding=16)
        file_box.grid(row=1, column=0, sticky="ew", pady=(10, 16))
        file_box.columnconfigure(0, weight=1)
        ttk.Label(file_box, text="文件名", style="InfoLabel.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(file_box, textvariable=self.file_name, style="InfoValue.TLabel", wraplength=720, justify="left").grid(row=1, column=0, sticky="w", pady=(6, 0))

        ttk.Label(card, text="结果目录", style="CardTitle.TLabel").grid(row=2, column=0, sticky="w")

        output_box = ttk.Frame(card, style="Info.TFrame", padding=16)
        output_box.grid(row=3, column=0, sticky="ew", pady=(10, 22))
        output_box.columnconfigure(0, weight=1)
        ttk.Label(output_box, text="目录路径", style="InfoLabel.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(output_box, textvariable=self.output_dir_var, style="PathValue.TLabel", wraplength=720, justify="left").grid(row=1, column=0, sticky="w", pady=(6, 0))

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=4, column=0, sticky="ew")
        actions.columnconfigure((0, 1, 2), weight=1)

        self.select_button = ttk.Button(actions, text="选择文件", style="Secondary.TButton", command=self.select_file)
        self.select_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.run_button = ttk.Button(actions, text="开始清理", style="Primary.TButton", command=self.run_cleanup)
        self.run_button.grid(row=0, column=1, sticky="ew", padx=8)

        self.open_dir_button = ttk.Button(actions, text="打开清理后的文件目录", style="Ghost.TButton", command=self.open_output_dir)
        self.open_dir_button.grid(row=0, column=2, sticky="ew", padx=(8, 0))

        self.status_card = ttk.Frame(container, style="Idle.TFrame", padding=18)
        self.status_card.grid(row=2, column=0, sticky="ew", pady=(18, 0))
        self.status_card.columnconfigure(0, weight=1)

        self.status_title_label = ttk.Label(self.status_card, textvariable=self.status_title, style="StatusTitle.TLabel")
        self.status_title_label.grid(row=0, column=0, sticky="w")
        self.status_body_label = ttk.Label(
            self.status_card,
            textvariable=self.status_body,
            style="StatusBody.TLabel",
            wraplength=760,
            justify="left",
        )
        self.status_body_label.grid(row=1, column=0, sticky="w", pady=(6, 0))

    def select_file(self) -> None:
        chosen = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not chosen:
            return

        resolved = Path(chosen).resolve()
        self.selected_file.set(str(resolved))
        self.file_name.set(resolved.name)
        self.output_dir = resolved.parent / "cleaned_output"
        self.output_dir_var.set(str(self.output_dir))
        self.last_output_file = None
        self.run_status = "ready"
        self.status_title.set("文件已就绪")
        self.status_body.set("确认无误后，点击开始清理。")
        self._update_run_button_state()
        self._refresh_status_card()

    def run_cleanup(self) -> None:
        selected = self.selected_file.get().strip()
        if not selected:
            messagebox.showwarning("请先选择文件", "请先选择一个 .xlsx 文件。")
            return

        self.run_status = "running"
        self.status_title.set("正在清理，请稍候")
        self.status_body.set("正在生成清理后的文件。")
        self._update_run_button_state()
        self._refresh_status_card()
        self.root.update_idletasks()

        try:
            report = clean_workbook(selected)
            self.output_dir = report.output_file.parent
            self.last_output_file = report.output_file
            self.output_dir_var.set(str(report.output_file.parent))
            self.run_status = "success"
            self.status_title.set("清理完成")
            self.status_body.set("新的文件已经生成，现在可以直接打开结果目录。")
            self._update_run_button_state()
            self._refresh_status_card()
            messagebox.showinfo("处理完成", f"已生成新文件：\n{report.output_file}")
        except CleanupRuleError as exc:
            self.last_output_file = None
            self.run_status = "error"
            self.status_title.set("清理未完成")
            self.status_body.set(str(exc))
            self._update_run_button_state()
            self._refresh_status_card()
            messagebox.showerror("处理失败", str(exc))
        except Exception as exc:  # pragma: no cover
            self.last_output_file = None
            self.run_status = "error"
            details = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            self.status_title.set("清理未完成")
            self.status_body.set(f"文件中存在需要先修正的问题。{details}")
            self._update_run_button_state()
            self._refresh_status_card()
            messagebox.showerror("处理失败", details)

    def open_output_dir(self) -> None:
        directory = self.output_dir
        if not directory or not directory.exists():
            messagebox.showinfo("没有结果目录", "当前还没有可打开的结果目录，请先处理一个文件。")
            return
        self._open_path(directory)

    def _update_run_button_state(self) -> None:
        has_file = bool(self.selected_file.get().strip())
        self.select_button.state(["!disabled"])

        if self.run_status == "running":
            self.run_button.configure(text="正在清理")
            self.run_button.state(["disabled"])
        else:
            self.run_button.configure(text="开始清理")
            if has_file:
                self.run_button.state(["!disabled"])
            else:
                self.run_button.state(["disabled"])

        if self.output_dir and self.output_dir.exists():
            self.open_dir_button.state(["!disabled"])
        else:
            self.open_dir_button.state(["disabled"])

    def _refresh_status_card(self) -> None:
        style_name = {
            "idle": "Idle.TFrame",
            "ready": "Idle.TFrame",
            "running": "Running.TFrame",
            "success": "Success.TFrame",
            "error": "Error.TFrame",
        }.get(self.run_status, "Idle.TFrame")
        self.status_card.configure(style=style_name)

    def _open_path(self, path: Path) -> None:
        if sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        elif os.name == "nt":
            os.startfile(str(path))
        else:
            subprocess.Popen(["xdg-open", str(path)])


def main() -> None:
    root = tk.Tk()
    ExcelCleanerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
