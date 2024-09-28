import time
from datetime import datetime, timedelta
from typing import Callable, Dict

from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer

from config import TARGET_DIR


class TriggerPipelineOnInsert(FileSystemEventHandler):
    def __init__(self, callback: Callable, args: Dict):
        self.callback = callback
        self.args = args
        self.last_modified = datetime.now()

    def on_modified(self, event: FileSystemEvent):
        """
        Check whether the desired file has been modified, and trigger the callback.
        Args:
            event (FileSystemEvent): Contains information about the file/directory modified.
        """
        if datetime.now() - self.last_modified < timedelta(seconds=1):
            return
        else:
            self.last_modified = datetime.now()
        filename = event.src_path
        encoded = filename.encode("utf-8").decode("utf-8")
        # Check for *any* excel file dumped inside
        if not event.is_directory and encoded.endswith(".xlsx"):
            self.args["EXPERT_LIST"] = encoded
            print("File modified!")
            self.callback(**self.args)


def init_fs_handler(callback: Callable, args: dict = {}):
    """
    Initialize the filesystem handler to run the `main()` function
    when a new file is inserted.
    Args:
        callback (Callable): Function to call when desired file modified.
        args (dict): callback arguments
    """
    start = time.time()
    fs_handler = TriggerPipelineOnInsert(callback, args)
    observer = Observer()
    observer.schedule(fs_handler, TARGET_DIR, recursive=True)
    observer.start()

    try:
        while True:
            observer.join(1)
    except KeyboardInterrupt:
        observer.stop()
    end = time.time()
    print(f"Finished after {end-start} secs.")


if __name__ == "__main__":
    init_fs_handler(lambda x: print(x))
