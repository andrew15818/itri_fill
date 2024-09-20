from typing import Callable, Dict
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler, FileSystemEvent


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
        #print(
        #    f"type: {event.event_type}\t  is directory: {event.is_directory}\t is synthetic: {event.is_synthetic}\t src path: {event.src_path}"
        #)
        if not event.is_directory and encoded == "./data\\iCAP會議套印資料_test.xlsx":
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
    init_fs_handler()
