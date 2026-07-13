from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OLD_PACKAGE = ROOT / "Company Package" / "銷售報表系統_公司版_20260620.zip"
NEW_PACKAGE = ROOT / "Company Package" / "銷售報表系統_公司版_20260713.zip"

REPLACEMENTS = {
    "銷售報表系統/scripts/daily_report_app.py": ROOT / "scripts" / "daily_report_app.py",
    "銷售報表系統/scripts/read_daily_orders.py": ROOT / "scripts" / "read_daily_orders.py",
    "銷售報表系統/scripts/report_pdfs.py": ROOT / "scripts" / "report_pdfs.py",
    "銷售報表系統/scripts/master_data.py": ROOT / "scripts" / "master_data.py",
    "銷售報表系統/config/app_settings.json": ROOT / "config" / "app_settings.json",
    "銷售報表系統/Order Template/order-file-template.xlsx": ROOT / "Order Template" / "order-file-template.xlsx",
    "銷售報表系統/Master Data/master-data.xlsx": ROOT / "Master Data" / "master-data.xlsx",
}

SKIP_PREFIXES = ("銷售報表系統/scripts/__pycache__/",)


def build_package() -> tuple[Path, int, str]:
    NEW_PACKAGE.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OLD_PACKAGE, "r") as source, zipfile.ZipFile(NEW_PACKAGE, "w", zipfile.ZIP_DEFLATED) as target:
        written: set[str] = set()
        for item in source.infolist():
            if item.filename in REPLACEMENTS:
                continue
            if item.filename in written:
                continue
            if item.filename.startswith(SKIP_PREFIXES):
                continue
            target.writestr(item, source.read(item.filename))
            written.add(item.filename)

        for archive_name, path in REPLACEMENTS.items():
            target.write(path, archive_name)
            written.add(archive_name)

    digest = hashlib.sha256(NEW_PACKAGE.read_bytes()).hexdigest()
    return NEW_PACKAGE, NEW_PACKAGE.stat().st_size, digest


def main() -> int:
    path, size, digest = build_package()
    print(path)
    print(size)
    print(digest)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
