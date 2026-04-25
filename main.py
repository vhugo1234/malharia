import os
import sys
import customtkinter as ctk
from gui.main_window import AppMainWindow

def main():
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
    
    app = AppMainWindow()
    app.mainloop()

if __name__ == "__main__":
    main()
