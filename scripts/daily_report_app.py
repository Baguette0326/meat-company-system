from __future__ import annotations

import os
import json
import re
import sys
import threading
import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from master_data import MASTER_DATA_PATH, ensure_master_data, refresh_all_order_templates
from read_daily_orders import generate_daily_report, generate_monthly_report


APP_TITLE = "銷售報表系統"
PROJECT_ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
CONFIG_PATH = PROJECT_ROOT / "config" / "app_settings.json"
DATE_RE = re.compile(r"(\d{8})")


def load_settings() -> dict[str, str]:
    if not CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def save_settings(settings: dict[str, str]) -> None:
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def setting_path(settings: dict[str, str], key: str, default: Path) -> Path:
    raw_value = settings.get(key, "").strip()
    if not raw_value:
        return default
    path = Path(raw_value)
    return path if path.is_absolute() else PROJECT_ROOT / path


def portable_path(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(PROJECT_ROOT.resolve()))
    except ValueError:
        return str(path)


def infer_report_date(folder: Path) -> str:
    folder_match = DATE_RE.search(folder.name)
    if folder_match:
        return folder_match.group(1)

    for workbook in sorted(folder.glob("*.xlsx")):
        if workbook.name.startswith("~$"):
            continue
        file_match = DATE_RE.search(workbook.name)
        if file_match:
            return file_match.group(1)

    return datetime.now().strftime("%Y%m%d")


def default_report_path(order_folder: Path, export_folder: Path | None) -> Path:
    report_date = infer_report_date(order_folder)
    target_folder = export_folder if export_folder else order_folder
    return target_folder / f"每日報表_{report_date}.xlsx"


def default_monthly_report_path(export_folder: Path | None, month: str | None = None) -> Path:
    target_folder = export_folder if export_folder else PROJECT_ROOT / "Monthly Reports"
    report_month = month if month else datetime.now().strftime("%Y%m")
    return target_folder / f"月報表_{report_month}.xlsx"


def today_order_folder(base_folder: Path) -> Path:
    return base_folder / datetime.now().strftime("%Y%m%d")


def ensure_upcoming_order_folders(base_folder: Path, days: int = 7) -> None:
    base_folder.mkdir(parents=True, exist_ok=True)
    today = datetime.now().date()
    for offset in range(days + 1):
        folder = base_folder / (today + timedelta(days=offset)).strftime("%Y%m%d")
        folder.mkdir(parents=True, exist_ok=True)


class DailyReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x680")
        self.minsize(980, 560)
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

        self.folder_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(value="請先確認今日訂單資料夾，然後按「產生今日報表」。")
        self.last_report: Path | None = None
        self.last_monthly_report: Path | None = None
        self.last_daily_pdf: Path | None = None
        self.last_monthly_pdf: Path | None = None
        self.settings = load_settings()

        self._configure_style()
        self._build_ui()
        ensure_master_data()
        self._load_defaults()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.configure("TButton", font=("Microsoft JhengHei UI", 12), padding=(14, 8))
        style.configure("TLabelframe.Label", font=("Microsoft JhengHei UI", 12, "bold"))

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=18)
        root.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(root, text=APP_TITLE, font=("Microsoft JhengHei UI", 20, "bold"))
        title.pack(anchor=tk.W)

        subtitle = ttk.Label(root, text="放入訂單 Excel，然後產生每日或每月報表。", font=("Microsoft JhengHei UI", 13))
        subtitle.pack(anchor=tk.W, pady=(4, 16))

        folder_row = ttk.Frame(root)
        folder_row.pack(fill=tk.X, pady=6)
        ttk.Label(folder_row, text="訂單 Excel 放這裡", width=18, font=("Microsoft JhengHei UI", 12, "bold")).pack(side=tk.LEFT)
        ttk.Entry(folder_row, textvariable=self.folder_var, font=("Microsoft JhengHei UI", 11)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Button(folder_row, text="更改", command=self.choose_folder).pack(side=tk.LEFT)
        folder_help = ttk.Label(root, text="平日不用改。按「打開今日訂單資料夾」，把訂單 Excel 放進去。", foreground="#475569", font=("Microsoft JhengHei UI", 10))
        folder_help.pack(anchor=tk.W, padx=(188, 0))

        output_row = ttk.Frame(root)
        output_row.pack(fill=tk.X, pady=(12, 6))
        ttk.Label(output_row, text="每日報表存在這裡", width=18, font=("Microsoft JhengHei UI", 12, "bold")).pack(side=tk.LEFT)
        ttk.Entry(output_row, textvariable=self.output_var, font=("Microsoft JhengHei UI", 11)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Button(output_row, text="更改", command=self.choose_output).pack(side=tk.LEFT)
        output_help = ttk.Label(root, text="檔名自動產生，例如：每日報表_20260618.xlsx。", foreground="#475569", font=("Microsoft JhengHei UI", 10))
        output_help.pack(anchor=tk.W, padx=(188, 0))

        reports_row = ttk.Frame(root)
        reports_row.pack(fill=tk.X, pady=(18, 8))

        daily_box = ttk.LabelFrame(reports_row, text="每日報表", padding=14)
        daily_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        ttk.Label(daily_box, text="今天完成入單後使用。", foreground="#475569", font=("Microsoft JhengHei UI", 10)).pack(anchor=tk.W, pady=(0, 10))
        ttk.Button(daily_box, text="1. 打開今日訂單資料夾", command=self.open_order_folder).pack(anchor=tk.W, pady=4)
        self.generate_button = ttk.Button(daily_box, text="2. 產生今日報表", command=self.generate_report)
        self.generate_button.pack(anchor=tk.W, pady=4)
        ttk.Button(daily_box, text="3. 打開今日報表", command=self.open_report).pack(anchor=tk.W, pady=4)
        ttk.Button(daily_box, text="4. 打開今日 PDF", command=self.open_daily_pdf).pack(anchor=tk.W, pady=4)

        monthly_box = ttk.LabelFrame(reports_row, text="月報表", padding=14)
        monthly_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))
        ttk.Label(monthly_box, text="查看本月客戶、貨品、分類走勢。", foreground="#475569", font=("Microsoft JhengHei UI", 10)).pack(anchor=tk.W, pady=(0, 10))
        self.monthly_button = ttk.Button(monthly_box, text="1. 產生本月月報", command=self.generate_monthly_report)
        self.monthly_button.pack(anchor=tk.W, pady=4)
        ttk.Button(monthly_box, text="2. 打開本月月報", command=self.open_monthly_report).pack(anchor=tk.W, pady=4)
        ttk.Button(monthly_box, text="3. 打開本月 PDF", command=self.open_monthly_pdf).pack(anchor=tk.W, pady=4)

        master_box = ttk.LabelFrame(root, text="貨品及客戶清單", padding=14)
        master_box.pack(fill=tk.X, pady=(10, 8))
        ttk.Label(
            master_box,
            text="新增貨品或客戶時，先打開清單修改，儲存並關閉 Excel，然後更新訂單模板。",
            foreground="#475569",
            font=("Microsoft JhengHei UI", 10),
        ).pack(anchor=tk.W, pady=(0, 8))
        ttk.Button(master_box, text="1. 打開貨品及客戶清單", command=self.open_master_data).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(master_box, text="2. 更新訂單", command=self.refresh_template).pack(side=tk.LEFT)

        status_box = ttk.LabelFrame(root, text="狀態", padding=12)
        status_box.pack(fill=tk.BOTH, expand=True, pady=(12, 0))
        self.status_label = ttk.Label(
            status_box,
            textvariable=self.status_var,
            justify=tk.LEFT,
            anchor=tk.NW,
            font=("Microsoft JhengHei UI", 12),
        )
        self.status_label.pack(fill=tk.BOTH, expand=True)

        note = ttk.Label(root, text="提示：更新報表前，請先關閉已開啟的報表 Excel。", foreground="#92400E", font=("Microsoft JhengHei UI", 11))
        note.pack(anchor=tk.W, pady=(10, 0))

    def choose_folder(self) -> None:
        initialdir = str(setting_path(self.settings, "orders_base_folder", PROJECT_ROOT / "Orders"))
        selected = filedialog.askdirectory(title="選擇訂單資料夾", initialdir=initialdir)
        if not selected:
            return
        selected_path = Path(selected)
        # If user selects a YYYYMMDD folder, remember its parent as base.
        # If user selects base Orders folder, auto-use today's child folder.
        if DATE_RE.fullmatch(selected_path.name):
            base_folder = selected_path.parent
            folder = selected_path
        else:
            base_folder = selected_path
            folder = today_order_folder(base_folder)
        ensure_upcoming_order_folders(base_folder)
        self.settings["orders_base_folder"] = portable_path(base_folder)
        self.folder_var.set(str(folder))
        export_folder = setting_path(self.settings, "export_folder", PROJECT_ROOT / "Daily Reports")
        default_output = default_report_path(folder, export_folder)
        self.output_var.set(str(default_output))
        self.status_var.set(f"訂單主資料夾：\n{base_folder}\n\n今日訂單資料夾：\n{folder}")

    def choose_output(self) -> None:
        initial = self.output_var.get()
        initialdir = str(Path(initial).parent) if initial else ""
        selected = filedialog.asksaveasfilename(
            title="選擇報表輸出位置",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialdir=initialdir,
            initialfile=Path(initial).name if initial else "daily_report.xlsx",
        )
        if selected:
            self.output_var.set(selected)

    def generate_report(self) -> None:
        folder = Path(self.folder_var.get().strip())
        output = Path(self.output_var.get().strip())
        today_text = datetime.now().strftime("%Y%m%d")

        if not folder.exists():
            if messagebox.askyesno(APP_TITLE, f"今日訂單資料夾不存在，是否建立？\n\n{folder}"):
                folder.mkdir(parents=True, exist_ok=True)
            else:
                return
        if not folder.is_dir():
            messagebox.showerror(APP_TITLE, "請選擇有效訂單資料夾。")
            return
        if DATE_RE.fullmatch(folder.name) and folder.name != today_text:
            proceed = messagebox.askyesno(
                APP_TITLE,
                "\n".join(
                    [
                        "你選擇的資料夾日期不是今天。",
                        "",
                        f"今天：{today_text}",
                        f"選擇：{folder.name}",
                        "",
                        "是否仍然繼續產生報表？",
                    ]
                ),
            )
            if not proceed:
                return
        if not output.name.lower().endswith(".xlsx"):
            messagebox.showerror(APP_TITLE, "報表輸出必須是 .xlsx 檔案。")
            return

        self._save_defaults(folder, output)

        self.generate_button.config(state=tk.DISABLED)
        self.status_var.set("正在讀取訂單檔案，請稍候...")

        thread = threading.Thread(target=self._generate_worker, args=(folder, output), daemon=True)
        thread.start()

    def _generate_worker(self, folder: Path, output: Path) -> None:
        try:
            file_count, sold_row_count, issue_count = generate_daily_report(folder, output)
        except PermissionError:
            self.after(
                0,
                self._generation_failed,
                "無法寫入報表。請先關閉已開啟的 Excel 報表，再重試。",
            )
            return
        except Exception as exc:  # noqa: BLE001 - surface unexpected app errors.
            self.after(0, self._generation_failed, f"產生報表失敗：\n{exc}")
            return

        self.after(0, self._generation_succeeded, output, file_count, sold_row_count, issue_count)

    def generate_monthly_report(self) -> None:
        folder = Path(self.folder_var.get().strip())
        if DATE_RE.fullmatch(folder.name):
            orders_base = folder.parent
            month = folder.name[:6]
        else:
            orders_base = setting_path(self.settings, "orders_base_folder", PROJECT_ROOT / "Orders")
            month = datetime.now().strftime("%Y%m")

        export_folder = setting_path(self.settings, "monthly_export_folder", PROJECT_ROOT / "Monthly Reports")
        output = default_monthly_report_path(export_folder, month)

        self.monthly_button.config(state=tk.DISABLED)
        self.status_var.set("正在產生本月月報，請稍候...")
        thread = threading.Thread(target=self._monthly_worker, args=(orders_base, month, output), daemon=True)
        thread.start()

    def _monthly_worker(self, orders_base: Path, month: str, output: Path) -> None:
        try:
            file_count, sold_row_count, issue_count = generate_monthly_report(orders_base, month, output)
        except PermissionError:
            self.after(0, self._monthly_failed, "無法寫入月報。請先關閉已開啟的月報 Excel，再重試。")
            return
        except Exception as exc:  # noqa: BLE001
            self.after(0, self._monthly_failed, f"產生月報失敗：\n{exc}")
            return

        self.after(0, self._monthly_succeeded, output, file_count, sold_row_count, issue_count)

    def _monthly_succeeded(self, output: Path, file_count: int, sold_row_count: int, issue_count: int) -> None:
        self.monthly_button.config(state=tk.NORMAL)
        self.last_monthly_report = output
        self.last_monthly_pdf = output.with_suffix(".pdf")
        self.status_var.set(
            "\n".join(
                [
                    "本月 Excel 及 PDF 月報完成。",
                    f"總訂單數量：{file_count}",
                    f"售出貨品記錄數：{sold_row_count}",
                    f"需要檢查的問題：{issue_count}",
                    f"月報：{output}",
                ]
            )
        )
        self.settings["monthly_export_folder"] = portable_path(output.parent)
        save_settings(self.settings)
        if issue_count:
            messagebox.showwarning(APP_TITLE, "月報已產生，但有問題需要查看「問題」分頁。")
        else:
            messagebox.showinfo(APP_TITLE, "月報已產生。")

    def _monthly_failed(self, message: str) -> None:
        self.monthly_button.config(state=tk.NORMAL)
        self.status_var.set(message)
        messagebox.showerror(APP_TITLE, message)

    def _generation_succeeded(self, output: Path, file_count: int, sold_row_count: int, issue_count: int) -> None:
        self.generate_button.config(state=tk.NORMAL)
        self.last_report = output
        self.last_daily_pdf = output.with_suffix(".pdf")
        self.status_var.set(
            "\n".join(
                [
                    "Excel 及 PDF 報表完成。",
                    f"總訂單數量：{file_count}",
                    f"售出貨品記錄數：{sold_row_count}",
                    f"需要檢查的問題：{issue_count}",
                    f"報表：{output}",
                ]
            )
        )

        if issue_count:
            messagebox.showwarning(APP_TITLE, "報表已產生，但有問題需要查看「問題」分頁。")
        else:
            messagebox.showinfo(APP_TITLE, "報表已產生。")

    def _generation_failed(self, message: str) -> None:
        self.generate_button.config(state=tk.NORMAL)
        self.status_var.set(message)
        messagebox.showerror(APP_TITLE, message)

    def open_report(self) -> None:
        path_text = self.output_var.get().strip()
        report = self.last_report or (Path(path_text) if path_text else None)
        if not report or not report.exists():
            messagebox.showerror(APP_TITLE, "找不到報表檔案。請先產生報表。")
            return
        os.startfile(report)

    def open_daily_pdf(self) -> None:
        path_text = self.output_var.get().strip()
        report = self.last_daily_pdf or (Path(path_text).with_suffix(".pdf") if path_text else None)
        if not report or not report.exists():
            messagebox.showerror(APP_TITLE, "找不到今日 PDF。請先產生今日報表。")
            return
        os.startfile(report)

    def open_monthly_report(self) -> None:
        folder = Path(self.folder_var.get().strip())
        if DATE_RE.fullmatch(folder.name):
            month = folder.name[:6]
        else:
            month = datetime.now().strftime("%Y%m")
        export_folder = setting_path(self.settings, "monthly_export_folder", PROJECT_ROOT / "Monthly Reports")
        report = self.last_monthly_report or default_monthly_report_path(export_folder, month)
        if not report.exists():
            messagebox.showerror(APP_TITLE, "找不到本月月報。請先產生月報。")
            return
        os.startfile(report)

    def open_monthly_pdf(self) -> None:
        folder = Path(self.folder_var.get().strip())
        month = folder.name[:6] if DATE_RE.fullmatch(folder.name) else datetime.now().strftime("%Y%m")
        export_folder = setting_path(self.settings, "monthly_export_folder", PROJECT_ROOT / "Monthly Reports")
        report = self.last_monthly_pdf or default_monthly_report_path(export_folder, month).with_suffix(".pdf")
        if not report.exists():
            messagebox.showerror(APP_TITLE, "找不到本月 PDF。請先產生本月月報。")
            return
        os.startfile(report)

    def open_master_data(self) -> None:
        ensure_master_data()
        os.startfile(MASTER_DATA_PATH)

    def refresh_template(self) -> None:
        try:
            refreshed_paths = refresh_all_order_templates(MASTER_DATA_PATH)
        except PermissionError:
            messagebox.showerror(APP_TITLE, "無法更新訂單模板。請先關閉已開啟的訂單模板 Excel。")
            return
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, f"更新訂單模板失敗：\n{exc}")
            return
        paths = "\n".join(str(path) for path in refreshed_paths)
        self.status_var.set(f"訂單模板已更新下拉選單：\n{paths}")
        messagebox.showinfo(APP_TITLE, "訂單模板已更新。")

    def open_order_folder(self) -> None:
        folder = Path(self.folder_var.get().strip())
        if not folder.exists():
            folder.mkdir(parents=True, exist_ok=True)
        os.startfile(folder)

    def _load_defaults(self) -> None:
        base = setting_path(self.settings, "orders_base_folder", PROJECT_ROOT / "Orders")
        if not base.exists() and base.is_absolute():
            base = PROJECT_ROOT / "Orders"
        export_path = setting_path(self.settings, "export_folder", PROJECT_ROOT / "Daily Reports")
        if not export_path.exists() and export_path.is_absolute():
            export_path = PROJECT_ROOT / "Daily Reports"

        ensure_upcoming_order_folders(base)
        folder = today_order_folder(base)
        self.folder_var.set(str(folder))
        self.output_var.set(str(default_report_path(folder, export_path)))
        self.status_var.set(f"已自動選擇今日訂單資料夾：\n{folder}")

    def _save_defaults(self, folder: Path, output: Path) -> None:
        if DATE_RE.fullmatch(folder.name):
            self.settings["orders_base_folder"] = portable_path(folder.parent)
        self.settings["import_folder"] = portable_path(folder)
        self.settings["export_folder"] = portable_path(output.parent)
        save_settings(self.settings)


def main() -> int:
    app = DailyReportApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
