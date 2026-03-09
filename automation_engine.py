import win32gui
import win32api
import win32con
import win32com.client
import time

# Constants
WM_CHAR = 0x0102
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
VK_RETURN = 0x0D
VK_BACK = 0x08
VK_HOME = 0x24

class AutomationEngine:
    # --- TESTING CONFIG ---
    # Para testing 1.0 y para producción 0.02#
    TEST_DELAY = 0.02

    @staticmethod
    def get_windows():
        def enum_handler(hwnd, titles):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title: titles.append(title)
        titles = []
        win32gui.EnumWindows(enum_handler, titles)
        return sorted(list(set(titles)))

    def force_focus(self, parent_hwnd):
        # Clean and simple, no ALT key menu bar trigger
        try:
            win32gui.ShowWindow(parent_hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(parent_hwnd)
            time.sleep(0.05)
        except: 
            pass

    def get_input_target(self, parent_hwnd):
        found = [parent_hwnd]
        def callback(hwnd, extra):
            c_name = win32gui.GetClassName(hwnd)
            if "Edit" in c_name or "RichEdit" in c_name:
                found[0] = hwnd
                return False
            return True
        win32gui.EnumChildWindows(parent_hwnd, callback, None)
        return found[0]

    def send_input(self, hwnd, text, clear_first=False, is_excel=False, should_stop=None):
        # --- EXCEL LOGIC ---
        if is_excel:
            excel = self.get_excel_app()
            if excel:
                try:
                    excel.ActiveCell.NumberFormat = "@" 
                    excel.ActiveCell.Value = "" 
                    
                    def rollback_excel():
                        excel.ActiveCell.Value = "" 
                        excel.SendKeys("~~", True)

                    written_chars = 0
                    for char in text:
                        if should_stop and should_stop():
                            rollback_excel()
                            return False 
                        
                        current_val = excel.ActiveCell.Value or ""
                        excel.ActiveCell.Value = current_val + char
                        written_chars += 1
                        time.sleep(self.TEST_DELAY)
                    
                    # ONE LAST CHECK BEFORE ENTER
                    if should_stop and should_stop():
                        rollback_excel()
                        return False 
                    
                    excel.SendKeys("~", True)
                    time.sleep(0.1) 
                    return True
                except: return False

        # --- SIRSI / UNIVERSAL LOGIC ---
        target_hwnd = self.get_input_target(hwnd)
        self.force_focus(hwnd)
        time.sleep(0.01)

        if clear_first:
            self.force_focus(hwnd)
            time.sleep(0.05)

            # CTRL + A
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
            win32api.keybd_event(ord('A'), 0, 0, 0)
            win32api.keybd_event(ord('A'), 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.15)

            # BACKSPACE
            win32api.keybd_event(VK_BACK, 0, 0, 0)
            win32api.keybd_event(VK_BACK, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.15)
       
        written_chars = 0
        
        def rollback_sirsi():
            for _ in range(written_chars):
                win32gui.SendMessage(target_hwnd, WM_KEYDOWN, VK_BACK, 0)
                time.sleep(0.02)
                win32gui.SendMessage(target_hwnd, WM_KEYUP, VK_BACK, 0)
                time.sleep(0.02)

        for char in text:
            if should_stop and should_stop():
                rollback_sirsi()
                return False

            win32gui.SendMessage(target_hwnd, WM_CHAR, ord(char), 0)
            written_chars += 1
            time.sleep(self.TEST_DELAY)

        # ONE LAST CHECK BEFORE ENTER
        if should_stop and should_stop():
            rollback_sirsi()
            return False

        # ONLY PRESS ENTER IF COMPLETE AND NOT STOPPED
        win32gui.SendMessage(target_hwnd, WM_KEYDOWN, VK_RETURN, 0)
        win32gui.SendMessage(target_hwnd, WM_KEYUP, VK_RETURN, 0)
       
        time.sleep(0.05)
        return True

    @staticmethod
    def get_excel_app():
        try: return win32com.client.GetActiveObject("Excel.Application")
        except: return None

    def dock_window_right(self, hwnd):
        try:
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            width = int(screen_width * 0.60)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, screen_width - width, 0, width, screen_height, win32con.SWP_SHOWWINDOW | win32con.SWP_NOOWNERZORDER)
        except: pass

    def dock_window_left(self, hwnd):
        try:
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            width = int(screen_width * 0.40)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, width, screen_height, win32con.SWP_SHOWWINDOW | win32con.SWP_NOOWNERZORDER)
        except: pass