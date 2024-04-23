import sys


class RedirectStdout:
    def __init__(self, scroll_text):
        self.scroll_text = scroll_text
        self.sdt_out = sys.stdout
        sys.stdout = self

    def write(self, message):
        self.sdt_out.write(f"{message}")
        self.scroll_text.insert("end", message)
        self.scroll_text.see("end")

    def flush(self):
        self.sdt_out.flush()

    def restore(self):
        sys.stdout = self.sdt_out
