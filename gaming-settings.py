from pywinauto.application import Application
from pywinauto import Desktop

def setBrightness(brightness=50):
    monitorian_app = Application(backend="uia").start("C:\\Users\\Admin\\AppData\\Local\\Microsoft\\WindowsApps\\10186emoacht.Monitorian_0q7myvhtpbc7w\Monitorian.exe")
    # nvidia_app = Application(backend="uia").start("C:\\Program Files\\WindowsApps\\NVIDIACorp.NVIDIAControlPanel_8.1.961.0_x64__56jybvy8sckqj\\nvcplui.exe")


if __name__ == "__main__":
    setBrightness(20)