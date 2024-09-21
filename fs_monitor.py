import os
from typing import Callable, Dict

from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer

from config import TARGET_DIR, TARGET_FILE


class TriggerPipelineOnInsert(FileSystemEventHandler):
    def __init__(self, callback: Callable, args: Dict):
        self.callback = callback
        self.args = args

    def on_modified(self, event: FileSystemEvent):
        """
        Check whether the desired file has been modified, and trigger the callback.
        Args:
            event (FileSystemEvent): Contains information about the file/directory modified.
        """
        filename = event.src_path
        encoded = filename.encode("utf-8").decode("utf-8")
        # TODO: Allow for different excel filenames?
        target_file = os.path.join(TARGET_DIR, TARGET_FILE)
        if not event.is_directory and encoded == target_file:
            print("File modified!")
            self.callback(**self.args)


def init_fs_handler(callback: Callable, args:dict = {}):
    """
    Initialize the filesystem handler to run the `main()` function
    when a new file is inserted.
    Args:
        callback (Callable): Function to call when desired file modified.
        args (dict): callback arguments
    """
    fs_handler = TriggerPipelineOnInsert(callback, args)
    observer = Observer()
    observer.schedule(fs_handler, "./data", recursive=True)
    observer.start()

    try:
        while True:
            observer.join(1)
    except KeyboardInterrupt:
        observer.stop()


if __name__ == "__main__":
    init_fs_handler(lambda x: print(x))
