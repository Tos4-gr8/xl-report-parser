import os
import xlwings as xw
from gui.main_window import ExcelCalcApp


def on_closing():    
    """Handles secure application closure without console errors."""
    # 1. Safely clean up Excel instances to prevent ghost processes in the OS
    try:
        xw.apps.cleanup() 
    except Exception:
        pass
    
    # 2. Instantly terminate the process at the OS level
    # This prevents Tkinter from executing pending 'after' loops, keeping the console clean
    os._exit(0)


if __name__ == "__main__":
    # Initialize and run the application
    app = ExcelCalcApp()
    
    # Bind the window close event directly to the handler
    app.protocol("WM_DELETE_WINDOW", on_closing)
    
    app.mainloop()
