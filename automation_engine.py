import win32gui
import win32api
import win32con
import win32com.client
import time

class AutomationEngine:
    @staticmethod
    def get_windows():
        def enum_handler(hwnd, lparam):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title: lparam.append(title)
        titles = []
        win32gui.EnumWindows(enum_handler, titles)
        return sorted(list(set(titles)))

    @staticmethod
    def dock_window_right(hwnd):
        screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
        target_width = int(screen_width * 0.10)
        
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 
                              screen_width - target_width, 0, 
                              target_width, screen_height, 
                              win32con.SWP_SHOWWINDOW)

    @staticmethod
    def get_excel_app():
        try:
            return win32com.client.GetActiveObject("Excel.Application")
        except:
            return None

    @staticmethod
    def send_input(hwnd, text, clear_first=False, is_excel=False):
        # 1. Asegurar el foco (Crucial para que las teclas no se pierdan)
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.15) 

        # 2. Lógica específica para EXCEL (Usando COM para formatear y borrar)
        if is_excel:
            excel = AutomationEngine.get_excel_app()
            if excel:
                try:
                    # Cambiar formato a Texto "@"
                    excel.Selection.NumberFormat = "@"
                    # Si el checkbox está marcado, borramos el contenido de la celda vía COM
                    if clear_first:
                        excel.Selection.ClearContents()
                    time.sleep(0.05)
                except Exception as e:
                    print(f"Excel COM Error: {e}")

        # 3. Lógica de borrado para programas NO-EXCEL (Workflow, etc)
        elif clear_first:
            # Usamos una secuencia más segura: Control + Home, luego Shift + End, luego Backspace
            # Esto evita activar menús accidentales de "Archivo"
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
            win32api.keybd_event(ord('A'), 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(ord('A'), 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.05)
            win32api.keybd_event(win32con.VK_BACK, 0, 0, 0)
            win32api.keybd_event(win32con.VK_BACK, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.05)

        # 4. Escritura de caracteres
        for char in text:
            vk = win32api.VkKeyScan(char)
            if vk == -1: continue # Saltar si no se reconoce el char
            
            shift = (vk >> 8) & 1
            vk_code = vk & 0xFF
            
            if shift: win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
            win32api.keybd_event(vk_code, 0, 0, 0)
            win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
            if shift: win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.01)
        
        # 5. Enter final
        win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
        win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)