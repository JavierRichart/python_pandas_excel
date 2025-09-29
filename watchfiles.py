from pathlib import Path
from time import sleep, time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import sys


def base_dir() -> Path:
    return Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


class XlsxHandler(FileSystemEventHandler):
    def __init__(self, settle: float = 1.5):
        super().__init__()
        self.result = None
        self.settle = settle

    def maybe_set(self, path: str):
        p = Path(path)
        if p.suffix.lower() == ".xlsx" and not p.name.startswith("~$"):
            try:
                s1 = p.stat().st_size
                sleep(self.settle)
                s2 = p.stat().st_size
                if s1 == s2:
                    self.result = p
            except FileNotFoundError:
                pass

    def on_created(self, event):
        if not event.is_directory:
            self.maybe_set(event.src_path)

    def on_moved(self, event):
        if not event.is_directory:
            self.maybe_set(event.dest_path)


def wait_for_new_xlsx(folder: Path, timeout: float = 120.0) -> Path | None:
    handler = XlsxHandler()
    obs = Observer()
    obs.schedule(handler, str(folder), recursive=False)
    obs.start()
    t0 = time()
    try:
        while time() - t0 < timeout and handler.result is None:
            sleep(0.2)
        return handler.result
    finally:
        obs.stop()
        obs.join()


if __name__ == "__main__":
    folder = base_dir()
    f = wait_for_new_xlsx(folder, timeout=180)
    print(f.name if f else "")
