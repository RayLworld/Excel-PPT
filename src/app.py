import tkinter as tk

try:
    from src.ui import PPTReplaceGUI
except ModuleNotFoundError:
    from ui import PPTReplaceGUI


def main():
    root = tk.Tk()
    app = PPTReplaceGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
