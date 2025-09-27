from pathlib import Path
import time
import sys


def base_dir()-> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def wait_new_file(
        folder: Path,
        timeout: float=120.0,
        poll_interval: float = 1.0,
        settle_time: float = 2.0
) -> Path | None:
    def file_set() -> set[Path]:
        return {p for p in folder.glob("*.xlsx") if not p.name.startswith("~$")}
    base_line = file_set()
    t0 = time.time()

    while time.time() - t0 < timeout:
        current = file_set()
        new_files = current - base_line
        if new_files:
            for p in sorted(new_files, key=lambda x: x.stat().st_mtime, reverse=True):
                size1 = p.stat().st_size
                time.sleep(settle_time)
                size2 = p.stat().st_size
                if size1 == size2:
                    return p
        time.sleep(poll_interval)

    return None


if __name__ == "__main__":
    folder = base_dir()
    f = wait_new_file(folder, timeout=300)